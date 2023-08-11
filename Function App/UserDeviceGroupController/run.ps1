# Input bindings are passed in via param block.
param($Timer)

# Functions
function Get-AuthToken {
    <#
    .SYNOPSIS
        Retrieve an access token for the Managed System Identity.
    
    .DESCRIPTION
        Retrieve an access token for the Managed System Identity.
    
    .NOTES
        Author:      Nickolaj Andersen
        Contact:     @NickolajA
        Created:     2021-06-07
        Updated:     2021-06-07
    
        Version history:
        1.0.0 - (2021-06-07) Function created
    #>
    Process {
        # Get Managed Service Identity details from the Azure Functions application settings
        $MSIEndpoint = $env:MSI_ENDPOINT
        $MSISecret = $env:MSI_SECRET

        # Define the required URI and token request params
        $APIVersion = "2017-09-01"
        $ResourceURI = "https://graph.microsoft.com"
        $AuthURI = $MSIEndpoint + "?resource=$($ResourceURI)&api-version=$($APIVersion)"

        # Call resource URI to retrieve access token as Managed Service Identity
        $Response = Invoke-RestMethod -Uri $AuthURI -Method "Get" -Headers @{ "Secret" = "$($MSISecret)" }

        # Construct authentication header to be returned from function
        $AuthenticationHeader = @{
            "Authorization" = "Bearer $($Response.access_token)"
            "ExpiresOn" = $Response.expires_on
        }

        # Handle return value
        return $AuthenticationHeader
    }
}

function Get-AzureTableEntity {
    param (
        [parameter(Mandatory = $true, HelpMessage = "Specify the Resource Group containing the Storage Account.")]
        [ValidateNotNull()]
        [string]$ResourceGroup,
        
        [parameter(Mandatory = $true, HelpMessage = "Specify the Storage Account name.")]
        [ValidateNotNull()]
        [string]$StorageAccountName,

        [parameter(Mandatory = $true, HelpMessage = "Specify the table name.")]
        [ValidateNotNull()]
        [string]$TableName,

        [parameter(Mandatory = $true, HelpMessage = "Specify the partition key.")]
        [ValidateNotNull()]
        [string]$PartitionKey
    )
    Process {
        # Get the table item
        $TableItem = Get-AzTableTable -resourceGroup $ResourceGroup -storageAccountName $StorageAccountName -TableName $TableName -ErrorAction "Stop" -Verbose:$false

        # Get all entities for specified partionkey
        $TableRows = Get-AzTableRow -Table $TableItem -PartitionKey $PartitionKey -ErrorAction "Stop" -Verbose:$false

        # Handle return value
        return $TableRows
    }
}

function Get-AzureADGroupMembership {
    param (
        [parameter(Mandatory = $true, HelpMessage = "Specify the Azure AD group ID.")]
        [ValidateNotNull()]
        [string]$GroupID
    )
    Process {
        # Retrieve all current memberships for specified group
        $AzureADGroupMembers = Invoke-MSGraphOperation -Get -APIVersion "v1.0" -Resource "groups/$($GroupID)/members" -ErrorAction "Stop"
        
        # Handle return value
        return $AzureADGroupMembers
    }
}

function Set-AzureADGroupMembership {
    param (
        [parameter(Mandatory = $true, HelpMessage = "Specify the Azure AD group ID.")]
        [ValidateNotNull()]
        [string]$GroupID,

        [parameter(Mandatory = $true, HelpMessage = "Pass a JSON converted string of a custom object containing all the new members.")]
        [ValidateNotNull()]
        [string]$Members
    )
    Process {
        # Set new members of specified group
        Invoke-MSGraphOperation -Patch -APIVersion "v1.0" -Resource "groups/$($GroupID)" -Body $Members -ContentType "application/json" -ErrorAction "Stop"
    }
}

function Remove-AzureADGroupMembership {
    param (
        [parameter(Mandatory = $true, HelpMessage = "Specify the Azure AD group ID.")]
        [ValidateNotNull()]
        [string]$GroupID,

        [parameter(Mandatory = $true, HelpMessage = "Specify the member object ID to be removed.")]
        [ValidateNotNull()]
        [string]$ObjectID
    )
    Process {
        # Set new members of specified group
        Invoke-MSGraphOperation -Delete -APIVersion "v1.0" -Resource "groups/$($GroupID)/members/$($ObjectID)/`$ref" -ErrorAction "Stop"
    }
}

# Version control
# 2022-09-20 - Version 1.0.0
# 2023-07-09 - Version 1.1.0
# 2023-08-09 - Version 1.2.0

# Import AzTable module
Write-Information -MessageData "Importing required modules"
Import-Module -Name "AzTable" -Verbose:$false

# Get auth token
$Global:AuthenticationHeader = Get-AuthToken

# Read application settings variables
$AzureResourceGroup = $env:AzureResourceGroup
$AzureStorageAccountName = $env:AzureStorageAccountName
$AzureStorageAccountTableName = $env:AzureStorageAccountTableName
$PartitionKey = $env:PartitionKey

try {
    # Read Azure Table Storage entities to be processed
    $TableEntities = Get-AzureTableEntity -ResourceGroup $AzureResourceGroup -StorageAccountName $AzureStorageAccountName -TableName $AzureStorageAccountTableName -PartitionKey $PartitionKey

    # Process each table entity for user to device group membership mapping
    if ($TableEntities -ne $null) {
        foreach ($TableEntity in $TableEntities) {
            Write-Information -MessageData "Processing user to device group mapping for entity: $($TableEntity.RowKey)"

            # Check if table entity is active, if not, skip current entity
            if ($TableEntity.State -eq "Active") {
                Write-Information -MessageData "- Current user group name: $($TableEntity.UserGroupName)"
                Write-Information -MessageData "- Current device group name: $($TableEntity.DeviceGroupName)"
                
                try {
                    # Read user group membership identities
                    $UserGroupMembers = Get-AzureADGroupMembership -GroupID $TableEntity.UserGroupID
                    $UserGroupMembersCount = ($UserGroupMembers | Measure-Object).Count

                    # Process each user identity from user group
                    if ($UserGroupMembers -ne $null) {
                        Write-Information -MessageData "- Current user group members count: $($UserGroupMembersCount)"
                        # Construct list to hold each device directory object reference
                        $DeviceDirectoryObjectList = New-Object -TypeName "System.Collections.Generic.List[object]"

                        # Get existing device members
                        $DeviceGroupCurrentMembers = Get-AzureADGroupMembership -GroupID $TableEntity.DeviceGroupID
                        
                        # Calculate the number of members in the device group depending on the return value
                        if ($DeviceGroupCurrentMembers.PSObject.Properties -match "value") {
                            $DeviceGroupCurrentMembersCount = ($DeviceGroupCurrentMembers.value | Measure-Object).Count
                        }
                        else {
                            $DeviceGroupCurrentMembersCount = ($DeviceGroupCurrentMembers | Measure-Object).Count
                        }
                        Write-Information -MessageData "- Current device group members count: $($DeviceGroupCurrentMembersCount)"

                        # Process each user identity and get devices for mapping
                        Write-Information -MessageData "- Retrieving registered devices for each user identity"

                        # Set batch size
                        $BatchSize = 20

                        # Construct concurrent bag object for Graph API responses of each batch
                        $UserRegisteredDevices = New-Object -TypeName "System.Collections.Generic.List[object]"

                        # Construct list for batch requests
                        $BatchRequestList = New-Object -TypeName "System.Collections.Generic.List[object]"

                        # Construct batch requests with size defined by $BatchSize
                        for ($i = 0; $i -lt $UserGroupMembers.Count; $i += $BatchSize) {
                            # Calculate end position of current batch
                            $EndPosition = $i + $BatchSize - 1
                            if ($EndPosition -ge $UserGroupMembers.Count) {
                                $EndPosition = $UserGroupMembers.Count
                            }

                            # Set current index for each batch request based of current count of $i (current count of array items)
                            $Index = $i

                            # Construct list object for current batch request
                            $CurrentBatchList = New-Object -TypeName "System.Collections.Generic.List[object]"

                            # Process each item in array from current index to end position and add to current batch list
                            foreach ($UserGroupMember in $UserGroupMembers[$i..($EndPosition)]) {
                                $BatchRequest = [PSCustomObject]@{
                                    "Id" = ++$Index
                                    "Method" = "GET"
                                    "Url" = "users/$($UserGroupMember.Id)/registeredDevices"
                                }
                                $CurrentBatchList.Add($BatchRequest)
                            }

                            # Construct current batch request object for Graph API call containing all batch requests defined in $CurrentBatchList
                            $BatchRequest = @{
                                "Method" = "Post"
                                "Uri" = 'https://graph.microsoft.com/beta/$batch'
                                "ContentType" = "application/json"
                                "Headers" = $Global:AuthenticationHeader
                                "ErrorAction" = "Stop"
                                "Body" = @{
                                    "requests" = $CurrentBatchList
                                } | ConvertTo-Json
                            }
                            $BatchRequestList.Add($BatchRequest)
                        }

                        # Call Graph API for each batch request
                        foreach ($BatchRequestItem in $BatchRequestList) {
                            $Responses = Invoke-RestMethod @BatchRequestItem
                            foreach ($ResponseItem in $Responses.responses) {
                                foreach ($ResponseValueItem in $ResponseItem.body.value) {
                                    if ($ResponseValueItem -ne $null) {
                                        $UserRegisteredDevices.Add($ResponseValueItem)
                                    }
                                }
                            }
                        }

                        # Filter registered devices by managed state, management authority and operating system
                        $UserRegisteredDevices = $UserRegisteredDevices | Where-Object { ($PSItem.operatingSystem -eq "Windows") -and ($PSItem.isManaged -eq $true) -and ($PSItem.managementType -eq "MDM") }

                        # Calculate count for user registered devices prior to any filtering
                        $UserRegisteredDevicesCount = ($UserRegisteredDevices | Measure-Object).Count
                        Write-Information -MessageData "- Retrieved registered devices count: $($UserRegisteredDevicesCount)"

                        # Filter registered devices for compliance if required
                        if ($TableEntity.IsCompliant -eq $true) {
                            $UserRegisteredDevices = $UserRegisteredDevices | Where-Object { $PSItem.isCompliant -eq $true }
                            
                            # Output non-compliant device count
                            Write-Information -MessageData "- Non-compliant device count: $($UserRegisteredDevicesCount - $UserRegisteredDevices.Count)"
                        }

                        # Filter registered devices for enrollment profile name if required
                        if ($TableEntity.PSObject.Properties -match "EnrollmentProfileName") {
                            if (-not([string]::IsNullOrEmpty($TableEntity.EnrollmentProfileName))) {
                                # Split enrollment profile names into array
                                $EnrollmentProfileNames = $TableEntity.EnrollmentProfileName.Split(";")
                                
                                # Construct enrollment profile names string for output and regular expression matching
                                $EnrollmentProfileNamesString = $EnrollmentProfileNames -join "|"
                                Write-Information -MessageData "- Filtering registered devices for enrollment profile name: $($EnrollmentProfileNames -join ", ")"

                                # Filter registered devices for enrollment profile name
                                $UserRegisteredDevices = $UserRegisteredDevices | Where-Object { $PSItem.enrollmentProfileName -match $EnrollmentProfileNamesString }

                                # Output non-matching enrollment profile name count
                                Write-Information -MessageData "- Non-matching enrollment profile name count: $($UserRegisteredDevicesCount - $UserRegisteredDevices.Count)"
                            }
                        }

                        # Process each user registered device and add to device directory object list
                        foreach ($UserRegisteredDevice in $UserRegisteredDevices) {
                            # Add device directory object reference to list
                            $PSObject = [PSCustomObject]@{
                                "ID" = $UserRegisteredDevice.id
                                "Uri" = "https://graph.microsoft.com/v1.0/directoryObjects/$($UserRegisteredDevice.id)"
                            }
                            $DeviceDirectoryObjectList.Add($PSObject)
                        }

                        # Construct custom object for device group membership update
                        $DeviceGroupMemberships = [PSCustomObject]@{
                            "members@odata.bind" = @()
                        }

                        # Calculate count for newly mapped device directory objects to be added as members of device group
                        $DeviceDirectoryObjectListCount = $DeviceDirectoryObjectList.Count
                        Write-Information -MessageData "- Mapped device directory objects count: $($DeviceDirectoryObjectListCount)"

                        # Process device group members update
                        if ($DeviceDirectoryObjectListCount -ge 1) {
                            # Determine if current device group contains any members, if not, skip device removal process
                            $PSObjectProperties = $DeviceGroupCurrentMembers.PSObject.Properties.Name
                            if (("value" -in $PSObjectProperties) -and ("@odata.context" -in $PSObjectProperties)) {
                                Write-Information -MessageData "- Current device group contains no members, only new members will be added"
                            }

                            # Construct reference and difference hash tables
                            # Reference is for discovered registered devices to users in the user group
                            # Difference is for current device group members
                            $DeviceGroupReferenceHash = @{}
                            $DeviceGroupDifferenceHash = @{}

                            # Add each device directory object reference to reference hash table
                            foreach ($DeviceDirectoryObjectListItem in $DeviceDirectoryObjectList) {
                                $DeviceGroupReferenceHash.Add($DeviceDirectoryObjectListItem.ID, $DeviceDirectoryObjectListItem.Uri)
                            }
                            
                            # Add each device group member reference to difference hash table
                            if ($DeviceGroupCurrentMembersCount -ge 1) {
                                foreach ($DeviceGroupCurrentMember in $DeviceGroupCurrentMembers) {
                                    $DeviceGroupDifferenceHash.Add($DeviceGroupCurrentMember.id, $null)
                                }
                            }
                            else {
                                $DeviceGroupDifferenceHash.Add("NoMembers", $null)
                            }

                            # Remove each device directory object reference from reference hash table if it exists in difference hash table (removes duplicates)
                            foreach ($ReferenceItem in $DeviceDirectoryObjectList) {
                                if ($DeviceGroupDifferenceHash.ContainsKey($ReferenceItem.ID)) {
                                    $DeviceGroupDifferenceHash.Remove($ReferenceItem.ID)
                                    $DeviceGroupReferenceHash.Remove($ReferenceItem.ID)
                                }
                            }

                            # Add new device group members if required
                            if ($DeviceGroupReferenceHash.Keys.Count -ge 1) {
                                Write-Information -MessageData "- Device group update is required, as '$($DeviceGroupReferenceHash.Keys.Count)' members are not part of the current device group"
                                
                                # Determine if device group membership update should be batched or not
                                if ($DeviceGroupReferenceHash.Keys.Count -gt 20) {
                                    # Initiate batching process for device group membership update to work around the limit of 20 directory objects per request
                                    Write-Information -MessageData "- Device group membership update requires batching since device group reference table contains '$($DeviceGroupReferenceHash.Keys.Count)' references"

                                    # Initiate device directory objects counter
                                    $ProcessedCount = 0
                                    
                                    do {
                                        # Select batch objects from device directory object list
                                        $BatchCurrentObjects = $DeviceGroupReferenceHash.Values | Select-Object -Skip $ProcessedCount -First 20

                                        # Update current processed device directory objects counter
                                        $ProcessedCount = $ProcessedCount + ($BatchCurrentObjects | Measure-Object).Count

                                        # Set device directory object id's for custom object
                                        $DeviceGroupMemberships."members@odata.bind" = $BatchCurrentObjects

                                        try {
                                            # Update device group membership for current table entity relation between user and device groups
                                            Write-Information -MessageData "- Updating device group memberships with overall progress: $($ProcessedCount) / $($DeviceGroupReferenceHash.Keys.Count)"
                                            $Response = Set-AzureADGroupMembership -GroupID $TableEntity.DeviceGroupID -Members ($DeviceGroupMemberships | ConvertTo-Json)
                                        }
                                        catch [System.Exception] {
                                            Write-Warning -Message "Failed to update device group memberships (with batching). Error message: $($_.Exception.Message)"
                                        }

                                        # Cleanup resources for optimized memory utilization
                                        Remove-Variable -Name "BatchCurrentObjects"
                                    }
                                    until ($ProcessedCount -ge $DeviceGroupReferenceHash.Keys.Count)
                                }
                                else {
                                    # Process without batching
                                    Write-Information -MessageData "- Device group membership update does not require batching, device group reference table contains '$($DeviceGroupReferenceHash.Keys.Count)' references"
            
                                    # Set device directory object id's for custom object
                                    $DeviceGroupMemberships."members@odata.bind" = $DeviceGroupReferenceHash.Values

                                    try {
                                        # Update device group membership for current table entity relation between user and device groups
                                        Write-Information -MessageData "- Updating device group memberships with count: $($DeviceGroupReferenceHash.Keys.Count)"
                                        $Response = Set-AzureADGroupMembership -GroupID $TableEntity.DeviceGroupID -Members ($DeviceGroupMemberships | ConvertTo-Json)
                                    }
                                    catch [System.Exception] {
                                        Write-Warning -Message "Failed to update device group memberships (without batching). Error message: $($_.Exception.Message)"
                                    }
                                }
                            }
                            else {
                                Write-Information -MessageData "- Device group update is not required, as all members are already part of the current device group"
                            }

                            # Remove existing device group members if required
                            if ($DeviceGroupDifferenceHash.ContainsKey("NoMembers")) {
                                Write-Information -MessageData "- Device group cleanup is not required, as it contains no members"
                            }
                            else {
                                if ($DeviceGroupDifferenceHash.Keys.Count -ge 1) {
                                    # Cleanup device group members if required
                                    Write-Information -MessageData "- Device group cleanup is required, as '$($DeviceGroupDifferenceHash.Keys.Count)' members are not part of the new device directory object list"
                                    Write-Information -MessageData "- Removing device group memberships with count: $($DeviceGroupDifferenceHash.Keys.Count)"
                                    foreach ($DeviceGroupDifferenceHashItem in $DeviceGroupDifferenceHash.Keys) {
                                        try {
                                            $Response = Remove-AzureADGroupMembership -GroupID $TableEntity.DeviceGroupID -ObjectID $DeviceGroupDifferenceHashItem -ErrorAction "Stop"
                                        }
                                        catch [System.Exception] {
                                            Write-Warning -Message "Failed to remove device group member '$($DeviceGroupDifferenceHashItem)' from '$($TableEntity.DeviceGroupName)' with error message: $($_.Exception.Message)"
                                        }
                                    }
                                }
                                else {
                                    Write-Information -MessageData "- Device group cleanup is not required, as all existing members are part of the new device directory object list"
                                }
                            }
                        }
                        else {
                            Write-Information -MessageData "- Device group update is not required, as device directory object list was empty"
                        }
                    }
                    else {
                        Write-Information -MessageData "- Table entity for '$($TableEntity.UserGroupName)' user group contains no members, process mapping will be skipped"
                    }
                }
                catch [System.Exception] {
                    Write-Warning -Message "Failed to retrieve group memberships for user group '$($TableEntity.UserGroupName)' with error message: $($_.Exception.Message)"
                }
            }
            else {
                Write-Information -MessageData "- Current user to device group mapping entity state is '$($TableEntity.State)' and will not be processed"
            }
        }
    }
    else {
        Write-Warning -Message "Empty list of table entities returned, ensure the correct table name and partition key used is specified"
    }
}
catch [System.Exception] {
    Write-Warning -Message "Failed to get Table entities with error message: $($_.Exception.Message)"
}