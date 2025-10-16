# Import the JSON file
$data = Get-Content -Path "ad_export.json" -Raw | ConvertFrom-Json

function Get-SafeDate {
    param($DateString)
    try {
        return [datetime]$DateString
    }
    catch {
        Write-Warning "Ogiltigt datum: $DateString"
        return $null
    }
}


# Make a variable for todays date
$now = Get-Date

# Create a hash table to store user data
$ad_export_users = @()

# Loop thru users and add them to $ad_export_users
foreach ($user in $data.users) {
    $lastLogon = Get-SafeDate($user.lastLogon)
    $lastPassSet = Get-SafeDate($user.passwordLastSet)
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

# Get inactive users hwo has not logged in for 30+ days
$inactiveUsers = $ad_export_users | Where-Object { $_.lastLogon -lt $now.AddDays(-30) }

# Sort computers after last login in descending order and filter out the 10 with longest time
$sortedComputersLastLogon = $data.computers | Sort-Object -Property lastLogon -Descending | Select-Object -first 10

# Get user account that expires within 30 days
$accountExpieringsSoon = $ad_export_users | Where-Object { $_.AccountExpires -lt $now.AddDays(30) }

# Get the computers who hasn't log on for 90+ days
$computersNotSeenForLong = $data.computers | Where-Object { $_.lastLogon -lt $now.AddDays(-30) }

$oldPass = $ad_export_users | Where-Object { $_.PasswordLastSet -lt $now.AddDays(-90) }

# Export to inactive_users.csv
$inactiveUsers | Select-Object samAccountName, displayName, Department, Site, LastLogon, DaysInactive, AccountExpires |
Export-Csv -Path "inactive_users.csv" -NoTypeInformation -Encoding UTF8

# Export to days_since_password_set.csv
$sorted_ad_export_users_lastPass | Select-Object samAccountName, displayName, Department, DaysSincePasswordSet, Enabled, PasswordNeverExpires |
Export-Csv -Path "days_since_password_set.csv" -NoTypeInformation -Encoding UTF8

# Create formated report
$report = @"
ACTIVE DIRECTORY AUDIT
$("=" * 40)
Genererad: $($now)
Domän: $($data.domain)
Exporterad: $($data.export_date)

Executive summary
$("-" * 40)
Critical: $($accountExpieringsSoon.count) konton förfaller inom 30 dagar
Varning: $($inactiveUsers.Count) användare har inte loggat in på 30+ dagar
Varning: $($computersNotSeenForLong.count) datorer har inte loggat in på 30+ dagar
Säkerhet: $($oldPass.count) användare har lösenord äldre än 90 dagar

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

# add usercount per department
foreach ($group in $users_per_department) {
    $report += "  {0,-10} {1,3} användare `n" -f $group.Name, $group.Count
}

$report += "`nDatorer som inte checkat in på längst tid `n"
$report += $("-" * 40) + "`n"
$report += "{0,-17} {1,-14} {2,-20} `n" -f "  Datornamn", "Site", "Senaste inloggning"

# Add computers who hasn't log on for a long time
foreach ($computer in $sortedComputersLastLogon) {
    $report += "  " + "{0,-15} {1,-14} {2,-20} `n" -f $computer.Name, $computer.site, $computer.lastLogon
}

# Export the report to text file
$report | Out-File -FilePath  "ad_audit_report.txt" -Encoding utf8