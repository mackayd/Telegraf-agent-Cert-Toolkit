[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

. "$PSScriptRoot\ADCSMetricLib.ps1"
[Console]::WriteLine((Get-ADCSMetricValue -MetricName 'Soonest60'))
exit 0
