<#
.SYNOPSIS
Finds stale user and group references in Microsoft Entra Conditional Access policies.

.DESCRIPTION
Connects to Microsoft Graph with Policy.Read.All and Directory.Read.All permissions,
then checks the included and excluded user and group IDs in all Conditional Access
policies. References that cannot be resolved through directoryObjects/getByIds are
reported as stale. Built-in values such as All, None, and GuestsOrExternalUsers are
excluded from the checks.

Displays a summary and a table of stale references, then exports the results to
ca-stale-references.csv in the current working directory. An existing report at
that path is overwritten. Conditional Access policies are not modified.

.EXAMPLE
PS> .\Find_Stale_CA_references.ps1

Connects to Microsoft Graph, checks all Conditional Access policies, and writes
the stale-reference report to the current working directory.

.EXAMPLE
PS> Import-Csv .\ca-stale-references.csv | Where-Object Field -eq 'ExcludeGroups'

After running the script, displays stale group references used in policy exclusions
from the exported report.

.NOTES
Requires the Microsoft Graph PowerShell SDK commands Connect-MgGraph,
Get-MgIdentityConditionalAccessPolicy, and Invoke-MgGraphRequest, and an account
authorized to read Conditional Access policies and directory objects.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Connect-MgGraph -Scopes 'Policy.Read.All','Directory.Read.All' -NoWelcome

$policies = Get-MgIdentityConditionalAccessPolicy -All
$skip = 'All','None','GuestsOrExternalUsers'

# Collect every referenced user/group ID per policy
$refs = foreach ($p in $policies) {
    $u = $p.Conditions.Users
    foreach ($pair in @(
        @('IncludeUsers',$u.IncludeUsers), @('ExcludeUsers',$u.ExcludeUsers),
        @('IncludeGroups',$u.IncludeGroups), @('ExcludeGroups',$u.ExcludeGroups))) {
        foreach ($id in @($pair[1])) {
            if ($id -and $id -notin $skip) {
                [pscustomobject]@{ Policy=$p.DisplayName; State=$p.State; Field=$pair[0]; Id=$id }
            }
        }
    }
}

# Resolve in batches of 1000
$ids = $refs.Id | Sort-Object -Unique
$found = @()
for ($i = 0; $i -lt $ids.Count; $i += 1000) {
    $batch = $ids[$i..([Math]::Min($i+999, $ids.Count-1))]
    $r = Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/directoryObjects/getByIds' `
         -Body (@{ ids = @($batch) } | ConvertTo-Json)
    $found += $r.value.id
}

$stale = $refs | Where-Object Id -notin $found
"Policies: $($policies.Count) | Refs: $($ids.Count) | Stale: $(@($stale).Count)"
$stale | Sort-Object Policy, Field | Format-Table -AutoSize
$stale | Export-Csv .\ca-stale-references.csv -NoTypeInformation
