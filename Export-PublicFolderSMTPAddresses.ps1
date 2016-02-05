Set-ADServerSettings -ViewEntireForest $true

Write-Output "Fetching all mail enabled public folders..."

$mpf = Get-MailPublicFolder -WarningAction SilentlyContinue -ResultSize Unlimited

$mpfc = $mpf | Measure-Object

Write-Host "Analyzing $($mpfc.Count) public folders"
$i = 1
$pfError = 0
$pfSMTP = 0
$PFCol = @()

Foreach ($pf in $mpf) {

    Write-Progress -Activity "Checking public folder $($pf.Name)" -Status "Adding properties" -PercentComplete(($i/$mpfc.Count)*100)
    
    foreach($address in $pf.EmailAddresses) {
        if([string]$address.Prefix -eq "smtp") {
            
            $obj = New-Object System.Object
            $obj | Add-Member -type NoteProperty -name EmailAddress -value $address.AddressString
            $PFCol += $obj
            $pfSMTP++
        }
    }
        
    $i++    
}

Write-Progress -Activity "Exporting CSV" -Status "..." -PercentComplete 100

$PFCol | Sort-Object EmailAddress -Unique | Export-Csv -Path D:\SCRIPTS\MailEnabledPublicFoldersSMTPAddresses.txt -NoTypeInformation -Encoding UTF8 -Delimiter "|" -Force

Write-Host "Public folder check finished!"
Write-Host "$($pfSMTP) SMTP addresses found!"