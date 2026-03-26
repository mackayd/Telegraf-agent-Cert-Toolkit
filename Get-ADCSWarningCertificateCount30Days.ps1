[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

. "$PSScriptRoot\ADCSMetricLib.ps1"
[Console]::WriteLine((Get-ADCSMetricValue -MetricName 'Warning30'))
exit 0
