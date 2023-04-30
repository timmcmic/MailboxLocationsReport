#Define user variables - please update as appropriate prior to running code.

[string]$outputFileName = "mailboxLocation.csv" #Define the CSV file name.
[string]$outputFilePath = "C:\temp\" #Define the output file path.

#Define code variables.

[int]$numberOfAuxArchives = 0 #Non zero value if one or more auto expanding archives exist.
[boolean]$hasPrimaryMailbox = $false #If the recipient has primary mailbox in Office 365.
[boolean]$hasComponentShard = $false #If the recipient has a componentshard in Office 365.
[boolean]$hasMainArchive = $false #If the recipient has a main archive in Office 365.
[boolean]$hasAuxArchive = $false #If the recipient has one of more auto expanding archives.
[boolean]$locationFound = $false #Determines if a location was found when querying mailboxes.

[string]$ComponentSharedString = "ComponentShared"
[string]$primaryMailboxString = "Primary"
[string]$mainArchiveString = "MainArchive"
[string]$auxArchiveString = "AuxArchive"

[array]$workingRecipients = @() #Holds all recipients in the organization
[array]$outputArray = @() #Holds the output of any recipient found.
[array]$workingLocations = @() #Holds any locations found for the particular mailbox.

[int]$recipientCounter = 1
[int]$totalRecipients = 0

[string]$userSelection = "0"

$fullOutputPath = $outputFilePath + $outputFileName

#There is the chance that an administrator only wants to gather a subset of information.  Provide the administrator with choices.


Write-Host "========================================="
Write-Host "         GENERATE MAILBOX REPORT"
Write-Host "========================================="
Write-Host ""
Write-Host "1 : All Recipients"
Write-Host ""
Write-Host "2 : Mailbox Enabled Recipients"
Write-Host ""
Write-Host "3 : Mailbox Enabled Recipients with Archives Enabled"
Write-Host ""
Write-Host "4 : All Recipients with Archives"
Write-Host ""
Write-Host "5 : Office 365 / Unified Groups Only"
Write-Host ""
Write-Host "6 : Guest Recipients"
Write-Host ""

$userChoice = Read-Host "Enter Selection"

#Capture all recipient objects in the organization.  This allows us to test for any componentShard or Office 365 groups which are not just mailboxes.

try {
    write-host "Gathering all Office 365 Recipients"

    switch ($userChoice)
    {
        1
        {
            write-host "Select All Recipients"
            $workingRecipients = get-recipient -recipientTypeDetails GroupMailbox,UserMailbox,MailUser,GuestMailUser -resultsize Unlimited -errorAction STOP | select-object externalDirectoryObjectID,primarySMTPAddress,RecipientType,RecipientTypeDetails
        }
        2
        {
            write-host "Mailbox Enabled Recipients"
            $workingRecipients = get-mailbox -resultsize unlimited -errorAction STOP | select-object externalDirectoryObjectID,primarySMTPAddress,RecipientType,RecipientTypeDetails
        }
        3
        {
            write-host "Mailbox Enabled Recipients with Archives Enabled"
            $workingRecipients = Get-Mailbox -ResultSize unlimited -Filter {archiveStatus -eq "Active"} -ErrorAction Stop | select-object externalDirectoryObjectID,primarySMTPAddress,RecipientType,RecipientTypeDetails
        }
        4
        {
            write-host "All Recipients with Archives"
            $workingRecipients = get-recipient -Filter {archiveStatus -eq "Active"} -resultsize Unlimited -errorAction STOP | select-object externalDirectoryObjectID,primarySMTPAddress,RecipientType,RecipientTypeDetails
        }
        5
        {
            write-host "Office 365 / Unified Groups Only"
            $workingRecipients = get-unifiedGroup -resultsize Unlimited -errorAction STOP | select-object externalDirectoryObjectID,primarySMTPAddress,RecipientType,RecipientTypeDetails
        }
        6
        {
            write-host "Guest Recipients"
            $workingRecipients = get-recipient -filter {RecipientTypeDetails -eq "GuestMailUser"} -resultsize unlimited -errorAction STOP | select-object externalDirectoryObjectID,primarySMTPAddress,RecipientType,RecipientTypeDetails
        }
        default
        {
            write-host "Default no selection - all recipients."
            $workingRecipients = get-recipient -recipientTypeDetails GroupMailbox,UserMailbox,MailUser,GuestMailUser -resultsize Unlimited -errorAction STOP | select-object externalDirectoryObjectID,primarySMTPAddress,RecipientType,RecipientTypeDetails
        }
    }
}
catch {
    write-host "Unable to obtain all recipients in Office 365."
    write-error $_
}

$totalRecipients = $workingRecipients.Count

#Iterate through each of the recipients and determine if there are any mailbox locations.

foreach ($recipient in $workingRecipients)
{
    #Rest the working variables for this recipient.

    $workingLocations = @() 
    $numberOfAuxArchives = 0
    $hasPrimaryMailbox = $false 
    $hasComponentShard = $false 
    $hasMainArchive = $false
    $hasAuxArchive = $false
    $locationfound = $false 

    write-host "Processing recipient number: "$recipientCounter" of total: "$totalRecipients
    $recipientCounter++
    write-host "Testing: "$recipient.externalDirectoryObjectID

    #First try to get the mailbox locations by user.  When doing so you will get a complete return of all locations.

    try {
        write-host "Testing locations by user."

        $workingLocations += get-mailboxLocation -user $recipient.externalDirectoryObjectID -errorAction STOP

        write-host "Testing locations by user successful" -ForegroundColor Green -BackgroundColor Yellow

        $locationFound = $true
    }
    catch {

        try {
            write-host "Unable to obtain locations by user.  Attempt identity (works for Office 365 Groups / None Users)" -ForegroundColor Red -BackgroundColor Yellow

            $workingLocations += get-mailboxLocation -identity $recipient.externalDirectoryObjectID -errorAction STOP

            write-host "Testing locations by identity successful." -ForegroundColor Green -BackgroundColor Yellow

            $locationFound = $true
        }
        catch {
            $workingLocations = @()
            write-host "Testing by location unsuccessful - object does not qualify for locations." -ForegroundColor Red -BackgroundColor Yellow
            write-host "Do not add user to the output array."
        }
    }

    #At this time we have gathered an object that has at least one or more locations.

    if ($workingLocations.count -gt 0)
    {
        write-host "Testing locations to generated output."

        foreach ($location in $workingLocations)
        {
            if ($location.MailboxLocationType -eq $primaryMailboxString)
            {
                write-host "Primary mailbox found." -ForegroundColor Blue -BackgroundColor Yellow
    
                $hasPrimaryMailbox = $true
            }
            elseif ($location.MailboxLocationType -eq $mainArchiveString)
            {
                write-host "Primary archive found." -ForegroundColor Blue -BackgroundColor Yellow
    
                $hasMainArchive = $true
            }
            elseif ($location.MailboxLocationType -eq $ComponentSharedString)
            {
                write-host "Component shared found." -ForegroundColor Blue -BackgroundColor Yellow
    
                $hasComponentShard = $true
            }
            elseif ($location.MailboxLocationType -eq $ComponentSharedString)
            {
                write-host "Auto expanding archive found." -ForegroundColor Blue -BackgroundColor Yellow
    
                $hasAuxArchive = $true
    
                $numberOfAuxArchives=$numberOfAuxArchives+1
            }
        }
    
        #At this point the location information has been generated.  Now we can generate the output object.
    
        if ($locationFound -eq $TRUE)
        {
            $functionObject = New-Object PSObject -Property @{
                ExternalDirectoryObjectID = $recipient.externalDirectoryObjectID  
                PrimarySMTPAddress = $recipient.primarySMTPAddress
                LocationCount = $workingLocations.count
                HasPrimaryMailbox = $hasPrimaryMailbox
                HasMainArchive = $hasMainArchive
                HasComponentShard = $hasComponentShard
                HasAuxArchive = $hasAuxArchive
                NumberOfAuxArchives = $numberOfAuxArchives
                RecipientType = $recipient.RecipientType
                RecipientTypeDetails = $recipient.RecipientTypeDetails
            }
    
            $functionObject = $functionObject | select-object ExternalDirectoryObjectID,PrimarySMTPAddress,LocationCount,HasPrimaryMailbox,HasMainArchive,HasComponentShard,HasAuxArchive,NumberOfAuxArchives,RecipientType,RecipientTypeDetails
        
            $outputArray += $functionObject
        }
    }
    else 
    {
        write-host "Location count is not greater than zero - although command did not failed no locations returned." -ForegroundColor Red -BackgroundColor Blue
    }
}

write-host "Concluded testing for locations - output array." -ForegroundColor Green -BackgroundColor Yellow

$outputArray | export-csv -Path $fullOutputPath