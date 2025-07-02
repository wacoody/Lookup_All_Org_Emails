Import-Module ActiveDirectory

$searchBases = @(
    "OU=Groups,DC=demo,DC=local",
    "OU=Resources,DC=demo,DC=local",
    "OU=Users,DC=demo,DC=local"
)

$results = foreach ($base in $searchBases) {
    Get-ADUser -SearchBase $base -Filter * -Properties mail, proxyAddresses | ForEach-Object {
        $user = $_
        $smtpAddresses = $user.proxyAddresses |
            Where-Object { ($_ -like "smtp:*" -or $_ -like "SMTP:*") -and ($_ -notlike "*onmicrosoft.com*") }

        if (![string]::IsNullOrWhiteSpace($user.mail) -or $smtpAddresses.Count -gt 0) {
            [PSCustomObject]@{
                SamAccountName = $user.SamAccountName
                PrimaryEmail   = $user.mail
                SMTPAddresses  = ($smtpAddresses -join '; ')
            }
        }
    }
}

$results | Export-Csv -Path "$env:userprofile\desktop\Filtered_SMTP_Emails.csv" -NoTypeInformation -Encoding UTF8
