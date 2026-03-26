[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

. "$PSScriptRoot\ADCSMetricLib.ps1"
[Console]::WriteLine((Get-ADCSMetricValue -MetricName 'Critical14'))
exit 0
