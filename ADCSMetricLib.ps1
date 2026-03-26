Set-StrictMode -Version Latest

$script:ADCSMetricLogPath = Join-Path -Path $PSScriptRoot -ChildPath 'ADCSMetricErrors.log'

function Write-ADCSMetricLog {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    try {
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        Add-Content -Path $script:ADCSMetricLogPath -Value "$timestamp`t$Message"
    }
    catch {
        # Intentionally swallow logging failures so the metric script can still return its fallback integer.
    }
}

function Get-CertutilPath {
    $cmd = Get-Command certutil.exe -ErrorAction SilentlyContinue
    if (-not $cmd) {
        throw 'certutil.exe was not found.'
    }
    return $cmd.Source
}

function Parse-CertutilDate {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    $formats = @(
        'dd/MM/yyyy HH:mm',
        'd/M/yyyy HH:mm',
        'dd/MM/yyyy H:mm',
        'd/M/yyyy H:mm',
        'dd/MM/yyyy HH:mm:ss',
        'd/M/yyyy HH:mm:ss',
        'MM/dd/yyyy HH:mm',
        'M/d/yyyy HH:mm',
        'MM/dd/yyyy H:mm',
        'M/d/yyyy H:mm',
        'MM/dd/yyyy HH:mm:ss',
        'M/d/yyyy HH:mm:ss'
    )

    foreach ($cultureName in @('en-GB', 'en-US')) {
        $culture = [System.Globalization.CultureInfo]::GetCultureInfo($cultureName)
        foreach ($format in $formats) {
            $parsed = [datetime]::MinValue
            if ([datetime]::TryParseExact(
                $Value,
                $format,
                $culture,
                [System.Globalization.DateTimeStyles]::AssumeLocal,
                [ref]$parsed
            )) {
                return $parsed
            }
        }
    }

    return [datetime]::Parse($Value)
}

function Invoke-CertutilCsvLocal {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Arguments,

        [Parameter(Mandatory = $true)]
        [string[]]$Headers
    )

    $certutil = Get-CertutilPath
    $raw = & $certutil @Arguments 2>&1

    if ($LASTEXITCODE -ne 0) {
        throw ((($raw | Out-String).Trim()) -replace "`r?`n", ' | ')
    }

    $lines = @(
        $raw |
        ForEach-Object { $_.ToString().Trim() } |
        Where-Object {
            $_ -and
            $_ -notmatch '^CertUtil:' -and
            $_ -notmatch '^Maximum Row Index:' -and
            $_ -notmatch '^\d+\s+Rows?$' -and
            $_ -notmatch '^\d+\s+Row Properties' -and
            $_ -notmatch '^\d+\s+Request Attributes' -and
            $_ -notmatch '^\d+\s+Certificate Extensions' -and
            $_ -notmatch '^\d+\s+Total Fields'
        }
    )

    if ($lines.Count -eq 0) {
        return @()
    }

    return @($lines | ConvertFrom-Csv -Header $Headers)
}

function Get-ADCSIssuedCertificateRows {
    return Invoke-CertutilCsvLocal -Arguments @(
        '-silent',
        '-restrict', 'Disposition=20',
        '-out', 'RequestID,RequesterName,CommonName,SerialNumber,NotAfter,CertificateTemplate',
        '-view', 'Log', 'csv'
    ) -Headers @(
        'RequestID',
        'RequesterName',
        'CommonName',
        'SerialNumber',
        'NotAfter',
        'CertificateTemplate'
    )
}

function Get-ADCSRevokedLookup {
    $rows = Invoke-CertutilCsvLocal -Arguments @(
        '-silent',
        '-out', 'RequestID,SerialNumber',
        '-view', 'Revoked', 'csv'
    ) -Headers @(
        'RequestID',
        'SerialNumber'
    )

    $lookup = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($row in $rows) {
        if ($row.RequestID) {
            [void]$lookup.Add("RID:$($row.RequestID)")
        }
        if ($row.SerialNumber) {
            [void]$lookup.Add("SER:$($row.SerialNumber)")
        }
    }

    return ,$lookup
}

function Get-ADCSExpiringCertificates {
    param(
        [int]$WindowDays = 60
    )

    $now = Get-Date
    $cutoff = $now.AddDays($WindowDays)
    $revokedLookup = Get-ADCSRevokedLookup
    $issuedRows = Get-ADCSIssuedCertificateRows

    $results = foreach ($row in $issuedRows) {
        if (-not $row.NotAfter) { continue }

        if ($row.RequestID -and $revokedLookup.Contains("RID:$($row.RequestID)")) { continue }
        if ($row.SerialNumber -and $revokedLookup.Contains("SER:$($row.SerialNumber)")) { continue }

        try {
            $notAfter = Parse-CertutilDate -Value $row.NotAfter
        }
        catch {
            continue
        }

        $daysToExpiry = [int][Math]::Floor(($notAfter - $now).TotalDays)

        if ($daysToExpiry -lt 0) { continue }
        if ($notAfter -gt $cutoff) { continue }

        [pscustomobject]@{
            RequestID           = [string]$row.RequestID
            RequesterName       = [string]$row.RequesterName
            CommonName          = [string]$row.CommonName
            SerialNumber        = [string]$row.SerialNumber
            NotAfter            = $notAfter
            CertificateTemplate = [string]$row.CertificateTemplate
            DaysToExpiry        = $daysToExpiry
        }
    }

    return @($results | Sort-Object DaysToExpiry, NotAfter, CommonName)
}

function Get-ADCSMetricValue {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('CollectionStatus', 'Expiring60', 'Warning30', 'Critical14', 'Soonest60')]
        [string]$MetricName
    )

    try {
        switch ($MetricName) {
            'CollectionStatus' {
                [void](Get-ADCSExpiringCertificates -WindowDays 60)
                return 1
            }
            'Expiring60' {
                return @(Get-ADCSExpiringCertificates -WindowDays 60).Count

            }
            'Warning30' {
                return @((Get-ADCSExpiringCertificates -WindowDays 60) | Where-Object { $_.DaysToExpiry -le 30 }).Count
            }
            'Critical14' {
                return @((Get-ADCSExpiringCertificates -WindowDays 60) | Where-Object { $_.DaysToExpiry -le 14 }).Count
            }
            'Soonest60' {
                $items = @(Get-ADCSExpiringCertificates -WindowDays 60)
                if ($items.Count -eq 0) {
                    return -1
                }
            return [int]$items[0].DaysToExpiry
            }
        }
    }
    catch {
        Write-ADCSMetricLog -Message "$MetricName failed: $($_.Exception.Message)"
        switch ($MetricName) {
            'CollectionStatus' { return 0 }
            'Soonest60' { return -2 }
            default { return -1 }
        }
    }
}
