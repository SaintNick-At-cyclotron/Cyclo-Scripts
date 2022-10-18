#variables

#$users = Import-Csv .csv
$groupmemberdir = "C:\Temp\"
$admin = ""
$rootFolder = 'C:\temp\CycloScript'

#$roomactions = $false


$i = 0
$ii = 0
$errors = 0

#Functions

function Find-preReqs {
    if (Test-Path -Path $rootFolder) {
    } else {
        new-item -type directory -path $rootFolder -Force
    }
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
}

function Open-Explorer {
    $fd = New-Object system.windows.forms.openfiledialog
    $fd.InitialDirectory = '%'
    $fd.MultiSelect = $true
    $fd.showdialog()
    $csv = $fd.filename
}

Function Get-Folder($initialDirectory="") {
    $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
    $foldername.Description = "Select a folder"
    $foldername.rootfolder = "MyComputer"
    $foldername.SelectedPath = $initialDirectory

    if($foldername.ShowDialog() -eq "OK")
    {
        $folder += $foldername.SelectedPath
    }
    return $folder
}

function get-groupmembers {
    Get-DistributionGroupMember -Identity $Group.Name | Select Name | export-csv "$($rootFolder)\$($Group.Name)-members.csv" -ErrorAction silentlycontinue
    Get-DistributionGroup -Identity $Group.Name | Select ManagedBy | export-csv "$($rootFolder)\$($Group.Name)-owners.csv" -ErrorAction silentlycontinue
}

function Convert-groups {
    Find-preReqs
    try 
    {
        $Askcsv = Read-Host -Prompt 'Do you have a list a groups that you want to gather information on? (Y/N)'
        if ($Askcsv -eq 'y') 
        {
            Open-Explorer
            foreach ($Group in $csv)
            {
                get-groupmembers

                $i++
                write-host "Total Progress: $i / $($csv.Count)" -ForegroundColor Green
            }
        }
        else 
        {
            $Askcsv = Read-Host -Prompt 'Do you want to get information for all groups? (Y/N)'
            if ($Askcsv -eq 'y')
            {
                $csv = Get-DistributionGroup | Select Name
                foreach ($Group in $csv)
                {
                    get-groupmembers
    
                    $i++
                    write-host "Total Progress: $i / $($csv.Count)" -ForegroundColor Green
                }
            }
        }    
        $AskGroupUpgrade = Read-Host -Prompt 'Would you like to upgrade a distribution group? (Y/N)'
        if ($AskGroupUpgrade -eq 'y')
        {
            $AskWhichGroup = Read-Host -Prompt 'Which group should be upgraded?'
            $seperator = "@"
            $GroupName = $AskWhichGroup.split($seperator)
            Remove-DistributionGroup $AskWhichGroup -Confirm:$false -ErrorAction silentlycontinue
            Start-Sleep -Seconds 5
            New-UnifiedGroup -DisplayName $GroupName[0] -PrimarySmtpAddress $AskWhichGroup -AutoSubscribeNewMembers -ErrorAction silentlycontinue
            $csv = import-csv "$($rootFolder)\$($GroupName[0])-members.csv"
            write-host "Adding Members..." -ForegroundColor Blue
            foreach ($member in $csv)
            {
                Add-UnifiedGroupLinks -Identity $AskWhichGroup -LinkType Members -Links $member.Name -ErrorAction silentlycontinue

                $i++
                write-host "Total Progress: $i / $($csv.Count)" -ForegroundColor Green
            }
            $ownercsv = import-csv "$($rootFolder)\$($GroupName[0])-owners.csv"
            write-host "Adding Owners..." -ForegroundColor Blue
            foreach ($owner in $ownercsv)
            {
                Add-UnifiedGroupLinks -Identity $AskWhichGroup -LinkType Owners -Links $owner.ManagedBy -ErrorAction silentlycontinue

                $ii++
                write-host "Total Progress: $ii / $($ownercsv.Count)" -ForegroundColor Green
            }
        }

    }
    catch [System.Object]
    { 
        Write-Output "Oh no,a failure is you! $($User.email), $_"
        $errors++
    }
}

function Update-Groups {
    Find-preReqs
    $Menu = [ordered]@{

        1 = 'Create Distribution List'
      
        2 = 'Create M365 Groups'
      
        3 = 'Delete M365 Groups'

        4 = 'Delete + Purge M365 Groups'

        5 = 'Add Group Members'
      
        }
        $Result = $Menu | Out-GridView -PassThru  -Title 'Make a  selection'

        Switch ($Result)  {

        {$Result.Name -eq 1} 
        {
            $csv = Open-Explorer
            foreach ($Group in $csv)
            {
                try {
                        write-host "Processing $($Group.email)" -ForegroundColor Yellow
                        New-DistributionGroup -Name $Group.name -PrimarySmtpAddress $Group.email -Confirm:$false -ErrorAction silentlycontinue

                        $i++
                        write-host "Total Progress: $i / $($csv.Count)" -ForegroundColor Green
                    }
                catch [System.Object]
                    { 
                        Write-Output "Oh no,a failure is you! $($Group.email), $_"
                        $errors++
                    }
            }
            
        }

        {$Result.Name -eq 2} 
        {
            $csv = Open-Explorer
            foreach ($Group in $csv)
            {
                try {
                        write-host "Adding M365 Group $($Group.email)" -ForegroundColor Yellow
                        New-UnifiedGroup -DisplayName $Group.name -Alias $Group.name -PrimarySmtpAddress $Group.email -AutoSubscribeNewMembers -Confirm:$false -ErrorAction silentlycontinue

                        $i++
                        write-host "Total Progress: $i / $($csv.Count)" -ForegroundColor Green
                    }
                catch [System.Object]
                    { 
                        Write-Output "Oh no,a failure is you! $($Group.email), $_"
                        $errors++
                    }
            }
        }

        {$Result.Name -eq 3} 
        {
            $csv = Open-Explorer
            foreach ($Group in $csv)
            {
                try {
                        write-host "Deleting M365 Group $($Group.email)" -ForegroundColor Yellow
                        Remove-UnifiedGroup -Identity $Group.email -confirm:$False | Out-Null

                        $i++
                        write-host "Total Progress: $i / $($csv.Count)" -ForegroundColor Green
                    }
                catch [System.Object]
                    { 
                        Write-Output "Oh no,a failure is you! $($Group.email), $_"
                        $errors++
                    }
            }
        }   
        {$Result.Name -eq 4} 
        {
            $csv = Open-Explorer
            foreach ($Group in $csv)
            {
                try {
                        write-host "Deleting M365 Group $($Group.email)" -ForegroundColor Yellow
                        Remove-UnifiedGroup -Identity $Group.email -confirm:$False | Out-Null
                        write-host "Purging M365 Group $($Group.email)" -ForegroundColor Yellow
                        $purgefilter = Get-AzureADMSDeletedGroup -Filter " DisplayName eq '$($Group.name)' "
                        Remove-AzureADMSDeletedDirectoryObject -Id $purgefilter.Id | Out-Null

                        $i++
                        write-host "Total Progress: $i / $($csv.Count)" -ForegroundColor Green
                    }
                catch [System.Object]
                    { 
                        Write-Output "Oh no,a failure is you! $($Group.email), $_"
                        $errors++
                    }
            }
        }   
        {$Result.Name -eq 5} 
        {
            $folder = Get-Folder
            $file = Get-ChildItem $folder -Filter *.csv
            foreach ($csv in $file)
            {
                try {
                        write-host "Updating memberships..." -ForegroundColor Yellow
                        $members = import-csv "$($folder)\$($csv.name)"
                        $ii = 0
                        foreach ($member in $members){
                            if ($member.status -eq "ACTIVE")
                            {
                                if ($grouptype -eq "dist")
                                {
                                    Add-DistributionGroupMember -Identity $User.email -Member $member.email -ErrorAction silentlycontinue
                                    if ($member.role -eq "OWNER")
                                    {
                                        Write-Host "$($member.email) is an owner"
                                        Set-DistributionGroup -Identity $User.email -ManagedBy $member.email -ErrorAction silentlycontinue
                                    }
                                }
                                if ($grouptype -eq "m365")
                                {
                                    Add-UnifiedGroupLinks -Identity $User.email -LinkType Members -Links $member.email -ErrorAction silentlycontinue
                                    if ($member.role -eq "OWNER")
                                    {
                                        Write-Host "$($member.email) is an owner"
                                        Add-UnifiedGroupLinks -Identity $User.email -LinkType Owners -Links $member.email -ErrorAction silentlycontinue
                                    }
                                }
                            }
                            elseif ($member.status -eq "") {
                                Write-Host "Creating Contact $($member.email)"
                                $seperator = "@"
                                $contactname = $member.email.split($seperator)
                                New-MailContact -Name $contactname[0] -ExternalEmailAddress $member.email -ErrorAction silentlycontinue
                                if ($grouptype -eq "dist")
                                {
                                    Add-DistributionGroupMember -Identity $User.email -Member $member.email -ErrorAction silentlycontinue
                                }
                                if ($grouptype -eq "m365")
                                {
                                    Add-UnifiedGroupLinks -Identity $User.email -LinkType Members -Links $member.email -ErrorAction silentlycontinue
                                }
                            }
                            $ii++
                            write-host "Membership Progress: $ii / $($members.Count)" -ForegroundColor Blue
                        }
                        $i++
                        write-host "Total Progress: $i / $($file.Count)" -ForegroundColor Green
                    }
                catch [System.Object]
                    { 
                        Write-Output "Oh no,a failure is you! $($csv.name), $_"
                        $errors++
                    }
            }
        }
    }
}

function Update-Rooms {
    Open-Explorer
    foreach ($Room in $csv)
    {
        if ($User.category -eq "Room") 
        {
            New-Mailbox -Name $User.name -Room -PrimarySmtpAddress $User.email
            Set-Mailbox $User.name -ResourceCapacity $User.capacity -Office $User.building
        }

        if ($User.category -eq "Equipment")
        {
            New-Mailbox -Name $User.name -Equipment -PrimarySmtpAddress $User.email
        }

        Set-CalendarProcessing $User.name -AutomateProcessing AutoAccept
        Add-DistributionGroupMember -Identity "$($User.building)" -Member $User.email
        Set-Place -Identity $User.name -Building $User.building -Floor $User.floor -FloorLabel $User.section -Label $User.description -City "All Cities" -CountryOrRegion "US" -State "CA"

        $i++
        write-host "Total Progress: $i / $($csv.Count)" -ForegroundColor Green
    }
}

$functions = @('Convert-groups', 'Update-Groups', 'Update-Rooms')
$functions | Out-GridView -PassThru  | Invoke-Expression 

Write-Host "Errors: $errors" -ForegroundColor Red