# Import the JSON file
$data = Get-Content -Path "ad_export.json" -Raw | ConvertFrom-Json

# Make a variable for todays date
$now = Get-Date

# Create variables to store data
$ad_export_users = @()

# Loop thru users and add them to $ad_export_users
foreach ($user in $data.users) {
    $lastLogon = [datetime]$user.lastLogon
    $lastPassSet = [datetime]$user.passwordLastSet
    $daysSinceLastLogon = ($now - $lastLogon).Days
    $daysSinceLastPass = ($now - $lastPassSet).Days

    $entry = @{
        SamAccountName       = $user.samAccountName
        DisplayName          = $user.displayName
        Department           = $user.department
        Site                 = $user.site
        LastLogon            = $user.lastLogon
        DaysInactive         = $daysSinceLastLogon
        AccountExpires       = $user.accountExpires
        PasswordLastSet      = $user.passwordLastSet
        DaysSincePasswordSet = $daysSinceLastPass
        PasswordNeverExpires = $user.passwordNeverExpires
        Enabled              = $user.enabled
    }

    # Convert to object
    $ad_export_users += New-Object -TypeName PSObject -Property $entry
}

# Sort $ad_export after days sence password set
$sorted_ad_export_users_lastPass = $ad_export_users | Sort-Object DaysSincePasswordSet -Descending


# Group users per department
$users_per_department = $data.users | Group-Object -Property department

# Get inactive users hwo has not logged in for 30 days
$inactiveUsers = $ad_export_users | Where-Object { $_.lastLogon -lt $now.AddDays(-30) }

# Sort computers after last login in descending order and filter out the 10 with longest time
$sortedComputers = $data.computers | Sort-Object -Property lastLogon -Descending | Select-Object -first 10

Write-Host $sortedComputers
# Get the 10 computers who has longest time since checkin



# Export to inactive_users.csv
$inactiveUsers | Select-Object samAccountName, displayName, Department, Site, LastLogon, DaysInactive, AccountExpires |
Export-Csv -Path "inactive_users.csv" -NoTypeInformation -Encoding UTF8

# Export to days_since_password_set.csv
$sorted_ad_export_users_lastPass | Select-Object samAccountName, displayName, Department, DaysSincePasswordSet, Enabled, PasswordNeverExpires |
Export-Csv -Path "days_since_password_set.csv" -NoTypeInformation -Encoding UTF8

# Create formated report
$report = @"

ACTIVE DIRECTORY AUDIT

======================
Rapporten genererad: $($now)
Domän: $($data.domain)
Exporterad: $($data.export_date)

Inaktiva användare:
$("-" * 40)

"@

$report += "{0,-22}  {1,-14} {2,10}  `n" -f "  Namn", "Användarnamn", "Senaste inloggningen"

# Add all inactive users to report
foreach ($user in $inactiveUsers) {
    $report += "  " + "{0,-20}  {1,-14} {2,10}  `n" -f $user.displayName, $user.samAccountName, $user.LastLogon #`n"
}



$report += "`nAntal användare per avdelning `n"
$report += $("-" * 40) + "`n"

foreach ($group in $users_per_department) {
    $report += "  {0,-10} {1,3} användare `n" -f $group.Name, $group.Count
}

$report += "`nDatorer som inte checkat in på längst tid `n"
$report += $("-" * 40) + "`n"
$report += "{0,-17} {1,-14} {2,-20} `n" -f "  Datornamn", "Site", "Senaste inloggning"

foreach ($computer in $sortedComputers) {
    $report += "  " + "{0,-15} {1,-14} {2,-20} `n" -f $computer.Name, $computer.site, $computer.lastLogon
}

# Export the report to text file
$report | Out-File -FilePath  "ad_audit_report.txt" -Encoding utf8