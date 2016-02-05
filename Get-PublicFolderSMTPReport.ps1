$File = "D:\SCRIPTS\MailEnabledPublicFoldersSMTPAddresses.txt"

$smtpAddresses = Import-Csv -Path $File

$max = $smtpAddresses.Count

$count = 1
$result = @()

$servers = @("SERVER1","SERVER2")

Write-Host "$($max) addresses to verify"

function Get-TrackingCount {
    param($ExchangeServer, $emailAddress)
    $messageLogEntriesCount = 0

    try {
        $messageLogEntries = $null
        $messageLogEntries = Get-MessageTrackingLog -Server $ExchangeServer -Recipients $emailAddress -EventId DELIVER -ResultSize Unlimited
        $messageLogEntriesCount = $messageLogEntries.Count
    }
    catch {}

    $messageLogEntriesCount
}

foreach($address in $smtpAddresses) {

    $smtpAddress = $address.EmailAddress

    $obj = New-Object System.Object
    $obj | Add-Member -type NoteProperty -name EmailAddress -value $smtpAddress

    foreach($server in $servers) {
        Write-Progress -Activity "Checking SMTP address $($smtpAddress) [$count/$($max)]" -Status "Fetching Tracking Log $($server)" -PercentComplete(($count/$max)*100)    

        $messages = Get-TrackingCount -ExchangeServer $server -emailAddress $smtpAddress
        
        $obj | Add-Member -type NoteProperty -name $($server) -value $messages
    }

    $result+=$obj

    $count++
}

Write-Progress -Activity "Exporting CSV" -Status "..." -PercentComplete 100

$result | Sort-Object EmailAddress | Export-Csv -Path D:\SCRIPTS\MailEnabledPublicFolderSMTPCount.txt -NoTypeInformation -Encoding UTF8 -Delimiter "|" -Force