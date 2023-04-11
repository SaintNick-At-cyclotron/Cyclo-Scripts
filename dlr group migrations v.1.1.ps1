function Check-Prerequisites {
    param (
        [string]$OutputDirectory = "C:\temp\group-exports"
    )

    # Ensure the output directory exists
    if (-not (Test-Path $OutputDirectory)) {
        New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
    }

    # Check if Microsoft.Graph module is installed, if not install it
    $GraphModule = Get-Module -ListAvailable -Name Microsoft.Graph
    if (-not $GraphModule) {
        Install-Module Microsoft.Graph -Scope CurrentUser -Force
    } else {
        # Check for updates and install the latest version
        $LatestGraphModule = Find-Module -Name Microsoft.Graph
        if ($LatestGraphModule.Version -gt $GraphModule.Version) {
            Update-Module -Name Microsoft.Graph -Force
        }
    }
}

function Show-GroupSourceMenu {
    param (
        [string]$Title = 'Choose Group Source'
    )

    Clear-Host
    Write-Host "================ $Title ================"

    Write-Host "1: Cloud Groups"
    Write-Host "2: On-Premise Groups"
    Write-Host "3: Other"
    Write-Host "Q: Quit"
}

function Show-OtherOptionsMenu {
    param (
        [string]$Title = 'Choose Other Option'
    )

    Clear-Host
    Write-Host "================ $Title ================"

    Write-Host "1: Contact Deletions"
    Write-Host "2: Check Pre-requisites"
    Write-Host "Q: Quit"
}

function Show-GroupTypeMenu {
    param (
        [string]$Title = 'Choose Group Type'
    )

    Clear-Host
    Write-Host "================ $Title ================"

    Write-Host "1: Microsoft 365 Groups"
    Write-Host "2: Teams Groups"
    Write-Host "3: Security Groups"
    Write-Host "4: Dynamic Azure AD Groups"
    Write-Host "5: Distribution Lists"
    Write-Host "6: Mail-enabled Security Groups"
    Write-Host "7: All Group Types"
    Write-Host "8: Specified Test Groups"
    Write-Host "Q: Quit"
}

function Export-M365Groups {
    param (
        [string]$OutputDirectory = "C:\temp\group-exports",
        [int]$GroupTypeSelection,
        [switch]$Prerequisites = $true
    )

    # Output hashtable for group membership
    $groupMembership = @{}

    if ($Prerequisites) {
        Check-Prerequisites -OutputDirectory $OutputDirectory
    }

    # Authenticate with Microsoft Graph
    Connect-MgGraph -Scopes "Group.Read.All", "User.Read.All"

    # Get all groups from Microsoft 365
    Import-Module Microsoft.Graph.Groups
    $allGroups = Get-MgGroup -Top 100

    # Filter groups based on the user's selection
    $M365Groups = switch ($GroupTypeSelection) {
        1 { $allGroups | Where-Object { $_.GroupTypes -contains "Unified" -and (Get-MgGroupTeam -GroupId $_.Id -ErrorAction SilentlyContinue) -eq $null } }
        2 { $allGroups | Where-Object { $_.GroupTypes -contains "Unified" -and (Get-MgGroupTeam -GroupId $_.Id -ErrorAction SilentlyContinue) -ne $null } }
        3 { $allGroups | Where-Object { $_.SecurityEnabled -eq $true -and $_.GroupTypes -eq $null } }
        4 { $allGroups | Where-Object { $_.GroupTypes -contains "DynamicMembership" } }
        5 { $allGroups | Where-Object { $_.MailEnabled -eq $true -and $_.SecurityEnabled -eq $false } }
        6 { $allGroups | Where-Object { $_.MailEnabled -eq $true -and $_.SecurityEnabled -eq $true } }
        7 { $allGroups }
        8 {
            $testGroupNames = @(
                "MS Patch Notification",
                "NLHQS Mkt",
                "UnitedKingdom SysEng-Alerts",
                "Local ICT Amsterdam",
                "Alerts Netherlands"
            )
            $allGroups | Where-Object { $_.DisplayName -in $testGroupNames }
        }
        default { $allGroups }
    }

    # Loop through each group
    foreach ($group in $M365Groups) {
        # Get group members
        $groupMembers = Get-MgGroupMember -GroupId $group.Id -Top 100

        # Get user details for each member
        $memberDetails = @()
        foreach ($member in $groupMembers) {
            if ($member['@odata.type'] -eq "#microsoft.graph.user") {
                $user = Get-MgUser -UserId $member.Id
                $userDetails = [pscustomobject]@{
                    Id            = $user.Id
                   DisplayName   = $user.DisplayName
                   GivenName     = $user.GivenName
                    Surname       = $user.Surname
                    UserPrincipalName = $user.UserPrincipalName
                    JobTitle      = $user.JobTitle
                    Department    = $user.Department
                    CompanyName   = $user.CompanyName
                   OfficeLocation = $user.OfficeLocation
               }
            } elseif ($member['@odata.type'] -eq "#microsoft.graph.group") {
                $group = Get-MgGroup -GroupId $member.Id
                $userDetails = [pscustomobject]@{
                    Id            = $group.Id
                    DisplayName   = $group.DisplayName
                    GroupTypes    = ($group.GroupTypes -join ';')
                    SecurityEnabled = $group.SecurityEnabled
                    Mail          = $group.Mail
                    MailNickname  = $group.MailNickname
                }
            } else {
                continue
            }
            $memberDetails += $userDetails
        }

        # Group details
        $groupDetails = [pscustomobject]@{
            Id               = $group.Id
            DisplayName      = $group.DisplayName
            Description      = $group.Description
            Mail             = $group.Mail
            MailEnabled      = $group.MailEnabled
            MailNickname     = $group.MailNickname
            SecurityEnabled  = $group.SecurityEnabled
            Visibility       = $group.Visibility
            GroupTypes       = ($group.GroupTypes -join ';')
            CreatedDateTime  = $group.CreatedDateTime
            Members          = ($memberDetails | ConvertTo-Json -Compress)
        }

        # Export group and member details to CSV
        $csvFilePath = Join-Path -Path $OutputDirectory -ChildPath $("group_" + $group.Id + "_details.csv")
        $groupDetails | Export-Csv -Path $csvFilePath -NoTypeInformation

        # Add group details to hashtable
        $groupMembership[$group.Id] = $groupDetails
    }

    # Disconnect from Microsoft Graph
    Disconnect-MgGraph

    # Return hashtable containing group membership
    $script:groupMembership = $groupMembership
}

function Export-OnPremGroups {
    param (
        [string]$OutputDirectory = "C:\temp\group-exports",
        [switch]$Prerequisites = $true
    )

    # Output hashtable for group membership
    $groupMembership = @{}

    if ($Prerequisites) {
        Check-Prerequisites -OutputDirectory $OutputDirectory
    }

    # Import the ActiveDirectory module
    Import-Module ActiveDirectory

    # Get On-Premise Mail Enabled Security Groups
    $OnPremGroups = Get-ADGroup -Filter {GroupCategory -eq "Security" -and mail -like "*"}

    # Loop through each group
    foreach ($group in $OnPremGroups) {
        # Get group members
        $groupMembers = Get-ADGroupMember -Identity $group -Recursive

        # Get user details for each member
        $memberDetails = @()
        foreach ($member in $groupMembers) {
            $memberObj = Get-ADObject -Identity $member -Properties DisplayName, Mail, SamAccountName, UserPrincipalName, DistinguishedName
            $memberDetails += [pscustomobject]@{
                Id                 = $member.ObjectGUID
                DisplayName        = $memberObj.DisplayName
                Mail               = $memberObj.Mail
                SamAccountName     = $memberObj.SamAccountName
                UserPrincipalName  = $memberObj.UserPrincipalName
                DistinguishedName  = $memberObj.DistinguishedName
            }
        }
        

        # Group details
        $groupDetails = [pscustomobject]@{
            Id               = $group.ObjectGUID
            DisplayName      = $group.Name
            Description      = $group.Description
            DistinguishedName = $group.DistinguishedName
            SamAccountName   = $group.SamAccountName
            GroupCategory    = $group.GroupCategory
            GroupScope       = $group.GroupScope
            Created          = $group.WhenCreated
            Changed          = $group.WhenChanged
            Members          = ($memberDetails | ConvertTo-Json -Compress)
        }

        # Export group and member details to CSV
        $csvFilePath = Join-Path -Path $OutputDirectory -ChildPath $("group_" + $group.ObjectGUID + "_details.csv")
        $groupDetails | Export-Csv -Path $csvFilePath -NoTypeInformation

        # Add group details to hashtable
        $groupMembership[$group.ObjectGUID] = $groupDetails
    }

    # Return hashtable containing group membership
    $script:groupMembership = $groupMembership
}

function ImportGroups {
    param (
        [string]$InputDirectory
    )

    # Authenticate with Microsoft Graph
    Connect-MgGraph -Scopes "Group.ReadWrite.All", "Directory.ReadWrite.All"

    $csvFiles = Get-ChildItem -Path $InputDirectory -Filter "group_*_details.csv"

    foreach ($csvFile in $csvFiles) {
        $GroupDetails = Import-Csv -Path $csvFile.FullName
        CreateOrUpdateGroups -GroupDetails $groupDetails
    }

    #Disconnect from MgGraph
    Disconnect-MgGraph
    
}

function CreateOrUpdateGroups {
    param (
        [pscustomobject]$GroupDetails
    )

    $existingGroup = Get-MgGroup -Filter "mail eq '$($GroupDetails.Mail)'" -ErrorAction Stop

    if ($existingGroup) {
        Write-Host "Group $($GroupDetails.DisplayName) already exists, skipping creation." -ForegroundColor Yellow
    }
    try {
        if ($GroupDetails.GroupTypes -contains "Unified") {
            # Microsoft 365 Group
            $newGroup = New-MgGroup -DisplayName $GroupDetails.DisplayName -Description $GroupDetails.Description -MailNickname $GroupDetails.MailNickname -Visibility $GroupDetails.Visibility -GroupTypes $GroupDetails.GroupTypes -MailEnabled:$true -SecurityEnabled:$false
        }
        elseif ($GroupDetails.SecurityEnabled -eq $true) {
            # Security Group
            $newGroup = New-MgGroup -DisplayName $GroupDetails.DisplayName -Description $GroupDetails.Description -MailNickname $GroupDetails.MailNickname -SecurityEnabled:$true
        }
        else {
            # Distribution Group
            $newGroup = New-MgGroup -DisplayName $GroupDetails.DisplayName -Description $GroupDetails.Description -MailNickname $GroupDetails.MailNickname
        }

        Write-Host "Group $($GroupDetails.DisplayName) created successfully." -ForegroundColor Green
        } catch {
            Write-Host "Failed to create group $($GroupDetails.DisplayName). Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Export-GroupMembership {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ExportFolder
    )

    # Authenticate with Microsoft Graph
    Connect-MgGraph -Scopes "User.Read.All", "Group.Read.All", "Contacts.Read"

    $allGroups = Get-MgGroup

    foreach ($group in $allGroups) {
        $exportData = @()
        $groupMembers = Get-MgGroupMember -GroupId $group.Id

        foreach ($member in $groupMembers) {
            $memberData = New-Object PSObject -Property @{
                Id   = $member.Id
                Type = $null
                Name = $null
            }

            $user = Get-MgUser -UserId $member.Id -ErrorAction SilentlyContinue
            if ($user) {
                $memberData.Type = 'User'
                $memberData.Name = $user.DisplayName
            } else {
                $contact = Get-MgContact -OrgContactId $member.Id -ErrorAction SilentlyContinue
                if ($contact) {
                    $memberData.Type = 'Contact'
                    $memberData.Name = $contact.DisplayName
                } else {
                    $memberData.Type = 'Unknown'
                }
            }

            $exportData += $memberData
        }

        $exportPath = Join-Path $ExportFolder "$($group.DisplayName)_members.csv"
        $exportData | Export-Csv -Path $exportPath -NoTypeInformation
    }
}

function Remove-MembersFromCSV {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ExportFolder,
        [Parameter(Mandatory = $true)]
        [string]$NamePrefix
    )

    # Helper function to delete a contact using direct HTTP request
    function Delete-Contact {
        param (
            [Parameter(Mandatory = $true)]
            [string]$OrgContactId
        )

        $graphApiVersion = "v1.0"
        $graphBaseUrl = "https://graph.microsoft.com/$graphApiVersion"
        $requestUrl = "$graphBaseUrl/contacts/$OrgContactId"

        Invoke-RestMethod -Uri $requestUrl -Headers $script:GraphToken -Method Delete
    }

    # Authenticate with Microsoft Graph and Exchange Online
    Connect-MgGraph -Scopes "User.Read.All", "Group.ReadWrite.All", "Contacts.ReadWrite"
    Connect-ExchangeOnline

    $csvFiles = Get-ChildItem -Path $ExportFolder -Filter *_members.csv

    # Store the access token for use in the helper function
    $script:GraphToken = @{
        "Authorization" = "Bearer " + (Get-MgContext).AccessToken
    }

    foreach ($csvFile in $csvFiles) {
        $groupDisplayName = $csvFile.BaseName.Replace("_members", "")
        $group = Get-MgGroup -Filter "DisplayName eq '$groupDisplayName'" | Select-Object -First 1

        if ($group) {
            $members = Import-Csv -Path $csvFile.FullName

            foreach ($member in $members) {
                if ($member.Name.StartsWith($NamePrefix)) {

                        # Try removing the member using Exchange Online PowerShell
                        $emailGroup = Get-DistributionGroup -Identity $group.DisplayName
                        if ($emailGroup) {
                            $recipient = Get-Recipient -Identity $member.Id
                            if ($recipient) {
                                Remove-DistributionGroupMember -Identity $emailGroup.PrimarySmtpAddress -Member $recipient.PrimarySmtpAddress -Confirm:$false
                            }
                    }

                    Write-Host "Removed $($member.Type) $($member.Name) from group $($group.DisplayName)" -ForegroundColor Green

                    if ($member.Type -eq 'Contact') {
                        # Delete the contact using the helper function
                        Delete-Contact -OrgContactId $member.Id
                        Write-Host "Deleted contact $($member.Name)" -ForegroundColor Green
                    }
                }
            }
        }
    }
}

function Remove-ContactsFromGroups {
    param (
        [Parameter(Mandatory = $true)]
        [string]$NamePrefix
    )

    # Authenticate with Microsoft Graph
    Connect-MgGraph -Scopes "User.Read.All", "Group.ReadWrite.All", "Contacts.ReadWrite"

    $allGroups = Get-MgGroup

    foreach ($group in $allGroups) {
        $groupMembers = Get-MgGroupMember -GroupId $group.Id

        foreach ($member in $groupMembers) {
            try {
                $user = Get-MgUser -UserId $member.Id
            } catch {
                $contact = Get-MgUserContact -UserId $member.Id
                if ($contact.DisplayName -like "$NamePrefix*") {
                    Remove-MgGroupMember -GroupId $group.Id -MemberId $member.Id -Confirm:$false
                    Write-Host "Removed contact $($contact.DisplayName) from group $($group.DisplayName)" -ForegroundColor Green

                    # Delete the contact
                    Remove-MgUserContact -UserId $member.Id -Confirm:$false
                    Write-Host "Deleted contact $($contact.DisplayName)" -ForegroundColor Green
                }
            }
        }
    }
}




# Parse command line arguments
$CommandLineArgs = @{
    Prerequisites = $true
    CreateGroups = $true
}
for ($i = 0; $i -lt $args.Length; $i++) {
    if ($args[$i] -eq "-prerequisites" -and ($i + 1) -lt $args.Length) {
        $CommandLineArgs.Prerequisites = [bool]::Parse($args[$i + 1])
        $i++
    } elseif ($args[$i] -eq "-creategroups" -and ($i + 1) -lt $args.Length) {
        $CommandLineArgs.CreateGroups = [bool]::Parse($args[$i + 1])
        $i++
    }
}

# Display the group source menu and get the user's selection
do {
    Show-GroupSourceMenu
    $groupSourceInput = Read-Host "Please make a selection"
    switch ($groupSourceInput) {
        '1' { $groupSource = 1; break }
        '2' { $groupSource = 2; break }
        '3' {
            do {
                Show-OtherOptionsMenu
                $otherOptionsInput = Read-Host "Please make a selection"
                switch ($otherOptionsInput) {
                    '1' { $otherOptions = 1; break }
                    '2' { $otherOptions = 2; break }
                    'q' { return }
                }
            } until ($otherOptionsInput -eq 'q' -or $otherOptions -ne $null)
            break
        }
        'q' { return }
    }
} until ($groupSourceInput -eq 'q' -or $groupSource -ne $null -or $otherOptions -ne $null)

if ($otherOptions -eq 1) {
    # Run contact deletions
    $NamePrefix = Read-Host "Enter the prefix to search for in contact names"
    Remove-ContactsFromGroups -NamePrefix $NamePrefix
} elseif ($otherOptions -eq 2) {
    # Check prerequisites
    Check-Prerequisites
}

if ($groupSource -eq 1) {
    # Display the group type menu and get the user's selection
    do {
        Show-GroupTypeMenu
        $groupTypeInput = Read-Host "Please make a selection"
        switch ($groupTypeInput) {
            '1' { $groupType = 1; break }
            '2' { $groupType = 2; break }
            '3' { $groupType = 3; break }
            '4' { $groupType = 4; break }
            '5' { $groupType = 5; break }
            '6' { $groupType = 6; break }
            '7' { $groupType = 7; break }
            '8' { $groupType = 8; break }
            'q' { return }
        }
    } until ($groupTypeInput -eq 'q' -or $groupType -ne $null)

    # Export M365 groups and their members based on the user's selection
    $exportedGroups = Export-M365Groups -Prerequisites:$CommandLineArgs.Prerequisites -GroupTypeSelection:$groupType

    # Create or update groups using the exported group membership
    $InputDirectory = "C:\temp\group-exports"
    ImportGroups -InputDirectory $InputDirectory
} elseif ($groupSource -eq 2) {
    # Export On-Premise groups and their members
    $exportedGroups = Export-OnPremGroups -Prerequisites:$CommandLineArgs.Prerequisites

    # Create or update groups using the exported group membership
    $InputDirectory = "C:\temp\group-exports"
    ImportGroups -InputDirectory $InputDirectory
}

