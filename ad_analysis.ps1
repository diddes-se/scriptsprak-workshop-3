# Import the JSON file
$data = Get-Content -Path "ad_export.json" -Raw | ConvertFrom-Json

# Create variables to stor data
$users_per_department = @{}

# Get inactive users hwo has not logged in for 30 days
$inactiveUsers = $data.users | Where-Object { 

    ([datetime]$_.lastLogon) -lt (Get-Date).AddDays(-30) 

}


foreach ($user in $data.users) {

    # Count users per department
    $dept = $user.department

    # If department dont exist in list create it
    if (-not $users_per_department.ContainsKey($dept)) {
        $users_per_department[$dept] = 0
    }

    # add 1 to the counter
    $users_per_department[$dept]++

}



# Create formated report
$report = @"

ACTIVE DIRECTORY AUDIT

======================

Domän: $($data.domain)
Exporterad: $($data.export_date)

Inaktiva användare:
$("-" * 40)
  Namn    Användarnamn   Sista inloggningen

"@

# Add all inactive users to report
foreach ($user in $inactiveUsers) {

    $report += "  $($user.displayName)`t$($user.samAccountName)`t$($user.lastLogon) `n"
}

$report += "`n Antal användare per avdelning `n"
$report += $("-" * 40) + "`n"

foreach ($dept in $users_per_department.Keys) {

    $report += "  ${dept}: $($users_per_department[$dept]) användare `n"

}



# Export the report to text file
$report | Out-File -FilePath  "ad_audit_report.txt" -Encoding utf8