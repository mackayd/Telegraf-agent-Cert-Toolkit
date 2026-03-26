[CmdletBinding()]
param(
    [int]$WindowDays = 60,
    [string]$OutputFolder = "$PSScriptRoot\Reports",
    [switch]$OpenReport
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

. "$PSScriptRoot\ADCSMetricLib.ps1"

function Convert-ToHtmlSafe {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return '' }
    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

try {
    if (-not (Test-Path -LiteralPath $OutputFolder)) {
        New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
    }

    $generatedAt = Get-Date
    $reportStamp = $generatedAt.ToString('yyyyMMdd-HHmmss')
    $reportBaseName = "ADCSExpiringCertificatesReport-$reportStamp"
    $htmlPath = Join-Path -Path $OutputFolder -ChildPath "$reportBaseName.html"
    $csvPath  = Join-Path -Path $OutputFolder -ChildPath "$reportBaseName.csv"
    $txtPath  = Join-Path -Path $OutputFolder -ChildPath 'ADCSExpiringCertificatesReport-latest.txt'

    $certs = @(Get-ADCSExpiringCertificates -WindowDays $WindowDays)

    $enriched = foreach ($cert in $certs) {
        $severity = if ($cert.DaysToExpiry -le 14) { 'critical' } elseif ($cert.DaysToExpiry -le 30) { 'warning' } else { 'notice' }
        [pscustomobject]@{
            RequestID           = [string]$cert.RequestID
            CommonName          = [string]$cert.CommonName
            RequesterName       = [string]$cert.RequesterName
            SerialNumber        = [string]$cert.SerialNumber
            CertificateTemplate = [string]$cert.CertificateTemplate
            NotAfter            = [datetime]$cert.NotAfter
            NotAfterDisplay     = ([datetime]$cert.NotAfter).ToString('yyyy-MM-dd HH:mm:ss')
            DaysToExpiry        = [int]$cert.DaysToExpiry
            Severity            = $severity
        }
    }

    $criticalCount = @($enriched | Where-Object { $_.Severity -eq 'critical' }).Count
    $warningCount  = @($enriched | Where-Object { $_.Severity -eq 'warning' }).Count
    $noticeCount   = @($enriched | Where-Object { $_.Severity -eq 'notice' }).Count
    $totalCount    = $enriched.Count
    $soonestDays   = if ($totalCount -gt 0) { ($enriched | Select-Object -First 1).DaysToExpiry } else { -1 }

    $csvRows = $enriched | Select-Object CommonName, RequesterName, RequestID, SerialNumber, CertificateTemplate, DaysToExpiry, Severity, NotAfterDisplay
    $csvRows | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

    $summaryText = @(
        "ADCS Expiring Certificate Report"
        "Generated: $($generatedAt.ToString('yyyy-MM-dd HH:mm:ss'))"
        "WindowDays: $WindowDays"
        "TotalExpiring: $totalCount"
        "Critical14Days: $criticalCount"
        "Warning30Days: $warningCount"
        "Notice31to${WindowDays}Days: $noticeCount"
        "SoonestDaysToExpiry: $soonestDays"
        "HTMLReport: $htmlPath"
        "CSVReport: $csvPath"
    )
    $summaryText | Set-Content -Path $txtPath -Encoding UTF8

    $chartGradient = if ($totalCount -gt 0) {
        $criticalPct = [math]::Round(($criticalCount / $totalCount) * 100, 2)
        $warningPct  = [math]::Round(($warningCount / $totalCount) * 100, 2)
        $noticePct   = 100 - $criticalPct - $warningPct
        "conic-gradient(#ef4444 0 ${criticalPct}%, #f59e0b ${criticalPct}% $(($criticalPct + $warningPct))%, #2563eb $(($criticalPct + $warningPct))% 100%)"
    }
    else {
        'conic-gradient(#334155 0 100%)'
    }

    $reportData = [pscustomobject]@{
        title = 'ADCS Expiring Certificate Report'
        generatedAt = $generatedAt.ToString('yyyy-MM-dd HH:mm:ss')
        windowDays = $WindowDays
        total = $totalCount
        critical = $criticalCount
        warning = $warningCount
        notice = $noticeCount
        soonestDays = $soonestDays
        csvFile = [System.IO.Path]::GetFileName($csvPath)
        certificates = @($enriched | ForEach-Object {
            [pscustomobject]@{
                requestId = $_.RequestID
                commonName = $_.CommonName
                requesterName = $_.RequesterName
                serialNumber = $_.SerialNumber
                certificateTemplate = $_.CertificateTemplate
                notAfter = $_.NotAfterDisplay
                daysToExpiry = $_.DaysToExpiry
                severity = $_.Severity
            }
        })
    }

    $json = $reportData | ConvertTo-Json -Depth 6 -Compress
    $json = $json -replace '</', '<\\/'

    $template = @'
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<title>__TITLE__</title>
<style>
:root {
  --bg: #071120;
  --panel: #0f1c33;
  --panel-soft: rgba(15, 28, 51, 0.88);
  --panel-2: #132442;
  --border: rgba(148, 163, 184, 0.16);
  --text: #e5eefc;
  --muted: #99aacd;
  --pass: #22c55e;
  --warn: #f59e0b;
  --info: #3b82f6;
  --critical: #ef4444;
  --shadow: 0 18px 40px rgba(0, 0, 0, 0.28);
}
* { box-sizing: border-box; }
body {
  margin: 0;
  font-family: "Segoe UI", Inter, Arial, sans-serif;
  background: radial-gradient(circle at top left, #0d2848 0%, #071120 42%, #050b16 100%);
  color: var(--text);
}
.container {
  width: min(1500px, calc(100% - 32px));
  margin: 16px auto 28px;
}
.hero {
  display: grid;
  grid-template-columns: 1.8fr 0.95fr;
  gap: 16px;
  background: linear-gradient(135deg, rgba(17, 31, 58, 0.96), rgba(8, 30, 53, 0.98));
  border: 1px solid var(--border);
  border-radius: 18px;
  padding: 18px;
  box-shadow: var(--shadow);
}
.hero h1 {
  margin: 0;
  font-size: 32px;
  line-height: 1.1;
}
.hero p {
  margin: 8px 0 0;
  color: var(--muted);
}
.hero-meta {
  display: flex;
  gap: 14px;
  flex-wrap: wrap;
  margin-top: 14px;
  color: var(--muted);
  font-size: 13px;
}
.badge {
  display: inline-flex;
  align-items: center;
  gap: 8px;
  padding: 6px 10px;
  border-radius: 999px;
  background: rgba(37, 99, 235, 0.18);
  border: 1px solid rgba(59, 130, 246, 0.24);
  color: #dbeafe;
  font-size: 12px;
  font-weight: 700;
  letter-spacing: 0.04em;
  text-transform: uppercase;
}
.summary-box {
  background: linear-gradient(180deg, rgba(17, 35, 61, 0.94), rgba(12, 25, 46, 0.92));
  border: 1px solid var(--border);
  border-radius: 16px;
  padding: 16px;
}
.summary-title {
  color: var(--muted);
  font-size: 12px;
  font-weight: 700;
  letter-spacing: 0.08em;
  text-transform: uppercase;
}
.bar-wrap {
  margin-top: 16px;
  background: #10203a;
  border-radius: 999px;
  overflow: hidden;
  height: 40px;
  display: flex;
}
.bar-segment {
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 12px;
  font-weight: 800;
  color: #f8fafc;
}
.bar-critical { background: #ef4444; }
.bar-warning { background: #f59e0b; }
.bar-notice  { background: #2563eb; }
.donut-card {
  display: flex;
  align-items: center;
  justify-content: center;
  background: linear-gradient(135deg, rgba(8, 26, 42, 0.85), rgba(13, 49, 44, 0.72));
  border: 1px solid var(--border);
  border-radius: 16px;
  min-height: 220px;
}
.donut {
  width: 164px;
  height: 164px;
  border-radius: 50%;
  background: __DONUT__;
  position: relative;
}
.donut::after {
  content: "";
  position: absolute;
  inset: 17px;
  border-radius: 50%;
  background: #0f1c33;
  border: 1px solid rgba(255,255,255,0.06);
}
.donut-center {
  position: absolute;
  inset: 0;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  z-index: 1;
  text-align: center;
}
.donut-number {
  font-size: 42px;
  font-weight: 800;
}
.donut-label {
  font-size: 12px;
  letter-spacing: 0.08em;
  text-transform: uppercase;
  color: var(--muted);
}
.cards {
  display: grid;
  grid-template-columns: repeat(5, minmax(0, 1fr));
  gap: 12px;
  margin-top: 14px;
}
.card {
  background: linear-gradient(180deg, rgba(15, 26, 46, 0.95), rgba(12, 22, 39, 0.92));
  border: 1px solid var(--border);
  border-radius: 14px;
  padding: 14px 16px;
  box-shadow: var(--shadow);
}
.card-label {
  color: var(--muted);
  font-size: 12px;
  text-transform: uppercase;
  letter-spacing: 0.08em;
}
.card-value {
  margin-top: 10px;
  font-size: 34px;
  font-weight: 800;
}
.layout {
  display: grid;
  grid-template-columns: 330px 1fr;
  gap: 16px;
  margin-top: 16px;
}
.panel {
  background: linear-gradient(180deg, rgba(14, 23, 40, 0.96), rgba(11, 19, 33, 0.94));
  border: 1px solid var(--border);
  border-radius: 16px;
  box-shadow: var(--shadow);
  overflow: hidden;
}
.panel-header {
  padding: 16px 18px 10px;
}
.panel-header h2 {
  margin: 0;
  font-size: 22px;
}
.panel-header p {
  margin: 6px 0 0;
  color: var(--muted);
  font-size: 13px;
}
.filter-row {
  display: flex;
  gap: 8px;
  flex-wrap: wrap;
  padding: 0 18px 14px;
}
.filter-btn {
  border: 1px solid rgba(59, 130, 246, 0.2);
  background: rgba(37, 99, 235, 0.08);
  color: #dbeafe;
  border-radius: 10px;
  padding: 8px 12px;
  font-weight: 700;
  cursor: pointer;
}
.filter-btn.active { outline: 2px solid rgba(59,130,246,0.35); }
.cert-list {
  max-height: 880px;
  overflow: auto;
  padding: 0 12px 14px;
}
.cert-item {
  padding: 12px;
  border: 1px solid rgba(59, 130, 246, 0.18);
  border-radius: 12px;
  margin-bottom: 10px;
  background: rgba(16, 29, 53, 0.88);
  cursor: pointer;
}
.cert-item.selected {
  border-color: #3b82f6;
  box-shadow: inset 0 0 0 1px rgba(59,130,246,0.45);
}
.cert-name {
  font-size: 16px;
  font-weight: 800;
  margin: 0;
}
.cert-sub {
  margin-top: 4px;
  color: var(--muted);
  font-size: 12px;
  word-break: break-word;
}
.pill {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  padding: 6px 10px;
  border-radius: 999px;
  font-size: 11px;
  font-weight: 800;
  text-transform: uppercase;
  letter-spacing: 0.08em;
}
.pill-critical { background: rgba(239,68,68,0.18); color: #fecaca; }
.pill-warning  { background: rgba(245,158,11,0.18); color: #fde68a; }
.pill-notice   { background: rgba(37,99,235,0.18); color: #bfdbfe; }
.detail-grid {
  display: grid;
  grid-template-columns: repeat(3, minmax(0,1fr));
  gap: 12px;
  padding: 0 18px 18px;
}
.detail-card {
  background: rgba(18, 30, 52, 0.9);
  border: 1px solid rgba(148, 163, 184, 0.14);
  border-radius: 12px;
  padding: 14px;
  min-height: 96px;
}
.detail-card .label {
  color: var(--muted);
  font-size: 11px;
  text-transform: uppercase;
  letter-spacing: 0.08em;
}
.detail-card .value {
  margin-top: 8px;
  font-size: 17px;
  font-weight: 700;
  word-break: break-word;
}
.table-wrap {
  margin: 0 18px 18px;
  border: 1px solid rgba(148, 163, 184, 0.14);
  border-radius: 12px;
  overflow: auto;
}
.table {
  width: 100%;
  border-collapse: collapse;
  min-width: 980px;
}
.table th,
.table td {
  padding: 12px 14px;
  border-bottom: 1px solid rgba(148,163,184,0.1);
  text-align: left;
  font-size: 13px;
}
.table th {
  position: sticky;
  top: 0;
  background: #10203a;
  color: #dbeafe;
  text-transform: uppercase;
  letter-spacing: 0.08em;
  font-size: 11px;
}
.table tbody tr:nth-child(odd) { background: rgba(14,26,47,0.7); }
.table tbody tr:hover { background: rgba(30,64,175,0.14); }
.footer-note {
  color: var(--muted);
  font-size: 12px;
  padding: 0 18px 18px;
}
.link {
  color: #93c5fd;
  text-decoration: none;
}
.empty {
  padding: 18px;
  color: var(--muted);
}
@media (max-width: 1200px) {
  .hero, .layout, .cards, .detail-grid { grid-template-columns: 1fr; }
}
</style>
</head>
<body>
<div class="container">
  <section class="hero">
    <div>
      <span class="badge">ADCS Local CA</span>
      <h1>__TITLE__</h1>
      <p>Interactive certificate expiry summary for the local Windows Certificate Authority.</p>
      <div class="hero-meta">
        <div>Generated at: __GENERATED__</div>
        <div>Window: __WINDOW__ days</div>
        <div>Total expiring: __TOTAL__</div>
        <div>Soonest: __SOONEST__</div>
      </div>
      <div class="summary-box" style="margin-top:16px;">
        <div class="summary-title">Expiry Severity Distribution</div>
        <div style="margin-top:6px;color:var(--muted);font-size:13px;">Click a certificate on the left to inspect the exact values that are driving the alert.</div>
        <div class="bar-wrap" aria-hidden="true">
          <div class="bar-segment bar-critical" style="width:__CRITICAL_WIDTH__%">CRITICAL __CRITICAL__</div>
          <div class="bar-segment bar-warning" style="width:__WARNING_WIDTH__%">WARN __WARNING__</div>
          <div class="bar-segment bar-notice" style="width:__NOTICE_WIDTH__%">NOTICE __NOTICE__</div>
        </div>
      </div>
    </div>
    <div class="donut-card">
      <div class="donut">
        <div class="donut-center">
          <div class="donut-number">__TOTAL__</div>
          <div class="donut-label">Expiring Certs</div>
        </div>
      </div>
    </div>
  </section>

  <section class="cards">
    <div class="card"><div class="card-label">Total Expiring</div><div class="card-value">__TOTAL__</div></div>
    <div class="card"><div class="card-label">Critical (≤14 days)</div><div class="card-value">__CRITICAL__</div></div>
    <div class="card"><div class="card-label">Warning (≤30 days)</div><div class="card-value">__WARNING__</div></div>
    <div class="card"><div class="card-label">Notice (31-__WINDOW__ days)</div><div class="card-value">__NOTICE__</div></div>
    <div class="card"><div class="card-label">Soonest Days To Expiry</div><div class="card-value">__SOONEST_RAW__</div></div>
  </section>

  <section class="layout">
    <aside class="panel">
      <div class="panel-header">
        <h2>Expiring Certificates</h2>
        <p>Select a certificate to inspect the values behind the alert.</p>
      </div>
      <div class="filter-row">
        <button class="filter-btn active" data-filter="all">All</button>
        <button class="filter-btn" data-filter="critical">Critical</button>
        <button class="filter-btn" data-filter="warning">Warning</button>
        <button class="filter-btn" data-filter="notice">Notice</button>
      </div>
      <div class="cert-list" id="certList"></div>
    </aside>

    <main class="panel">
      <div class="panel-header">
        <h2 id="detailTitle">Certificate Details</h2>
        <p id="detailSubtitle">Select a certificate from the list to view full details.</p>
      </div>
      <div class="detail-grid">
        <div class="detail-card"><div class="label">Status</div><div class="value" id="detailSeverity">-</div></div>
        <div class="detail-card"><div class="label">Days To Expiry</div><div class="value" id="detailDays">-</div></div>
        <div class="detail-card"><div class="label">Expiry Date</div><div class="value" id="detailNotAfter">-</div></div>
        <div class="detail-card"><div class="label">Requester</div><div class="value" id="detailRequester">-</div></div>
        <div class="detail-card"><div class="label">Request ID</div><div class="value" id="detailRequestId">-</div></div>
        <div class="detail-card"><div class="label">Serial Number</div><div class="value" id="detailSerial">-</div></div>
      </div>
      <div class="table-wrap">
        <table class="table">
          <tbody>
            <tr><th>Common Name</th><td id="detailCommonName">-</td></tr>
            <tr><th>Certificate Template</th><td id="detailTemplate">-</td></tr>
            <tr><th>Requester Name</th><td id="detailRequesterFull">-</td></tr>
          </tbody>
        </table>
      </div>
      <div class="panel-header" style="padding-top:0;">
        <h2>All Certificates In Window</h2>
        <p>CSV export: <a class="link" href="__CSV_FILE__">__CSV_FILE__</a></p>
      </div>
      <div class="table-wrap">
        <table class="table" id="allCertsTable">
          <thead>
            <tr>
              <th>Common Name</th>
              <th>Severity</th>
              <th>Days To Expiry</th>
              <th>Expiry Date</th>
              <th>Requester</th>
              <th>Request ID</th>
              <th>Template</th>
            </tr>
          </thead>
          <tbody id="allCertsBody"></tbody>
        </table>
      </div>
      <div class="footer-note">If this report shows a warning or critical count above zero, the corresponding Aria custom-script metrics should also reflect the same state.</div>
    </main>
  </section>
</div>
<script>
const reportData = __REPORT_JSON__;
const certList = document.getElementById('certList');
const allCertsBody = document.getElementById('allCertsBody');
const filterButtons = Array.from(document.querySelectorAll('.filter-btn'));
let currentFilter = 'all';
let selectedIndex = 0;

function severityClass(severity) {
  if (severity === 'critical') return 'pill pill-critical';
  if (severity === 'warning') return 'pill pill-warning';
  return 'pill pill-notice';
}

function setDetails(cert) {
  document.getElementById('detailTitle').textContent = cert ? cert.commonName : 'Certificate Details';
  document.getElementById('detailSubtitle').textContent = cert ? 'Inspect the exact certificate values behind the alert.' : 'Select a certificate from the list to view full details.';
  document.getElementById('detailSeverity').innerHTML = cert ? `<span class="${severityClass(cert.severity)}">${cert.severity}</span>` : '-';
  document.getElementById('detailDays').textContent = cert ? cert.daysToExpiry : '-';
  document.getElementById('detailNotAfter').textContent = cert ? cert.notAfter : '-';
  document.getElementById('detailRequester').textContent = cert ? cert.requesterName || '-' : '-';
  document.getElementById('detailRequestId').textContent = cert ? cert.requestId || '-' : '-';
  document.getElementById('detailSerial').textContent = cert ? cert.serialNumber || '-' : '-';
  document.getElementById('detailCommonName').textContent = cert ? cert.commonName || '-' : '-';
  document.getElementById('detailTemplate').textContent = cert ? cert.certificateTemplate || '-' : '-';
  document.getElementById('detailRequesterFull').textContent = cert ? cert.requesterName || '-' : '-';
}

function getFilteredCerts() {
  if (currentFilter === 'all') return reportData.certificates;
  return reportData.certificates.filter(c => c.severity === currentFilter);
}

function renderList() {
  const filtered = getFilteredCerts();
  if (selectedIndex >= filtered.length) selectedIndex = 0;
  certList.innerHTML = '';

  if (!filtered.length) {
    certList.innerHTML = '<div class="empty">No certificates match this filter.</div>';
    setDetails(null);
    return;
  }

  filtered.forEach((cert, idx) => {
    const item = document.createElement('div');
    item.className = 'cert-item' + (idx === selectedIndex ? ' selected' : '');
    item.innerHTML = `
      <div style="display:flex;align-items:center;justify-content:space-between;gap:10px;">
        <p class="cert-name">${cert.commonName || '(no common name)'}</p>
        <span class="${severityClass(cert.severity)}">${cert.severity}</span>
      </div>
      <div class="cert-sub">Days to expiry: ${cert.daysToExpiry}</div>
      <div class="cert-sub">Expiry: ${cert.notAfter}</div>
      <div class="cert-sub">Requester: ${cert.requesterName || '-'}</div>`;
    item.addEventListener('click', () => {
      selectedIndex = idx;
      renderList();
    });
    certList.appendChild(item);
  });

  setDetails(filtered[selectedIndex]);
}

function renderTable() {
  allCertsBody.innerHTML = '';
  reportData.certificates.forEach(cert => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${cert.commonName || '-'}</td>
      <td><span class="${severityClass(cert.severity)}">${cert.severity}</span></td>
      <td>${cert.daysToExpiry}</td>
      <td>${cert.notAfter}</td>
      <td>${cert.requesterName || '-'}</td>
      <td>${cert.requestId || '-'}</td>
      <td>${cert.certificateTemplate || '-'}</td>`;
    allCertsBody.appendChild(tr);
  });
  if (!reportData.certificates.length) {
    const tr = document.createElement('tr');
    tr.innerHTML = '<td colspan="7" class="empty">No expiring certificates found in the selected window.</td>';
    allCertsBody.appendChild(tr);
  }
}

filterButtons.forEach(btn => {
  btn.addEventListener('click', () => {
    filterButtons.forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    currentFilter = btn.dataset.filter;
    selectedIndex = 0;
    renderList();
  });
});

renderList();
renderTable();
</script>
</body>
</html>
'@

    $totalForWidth = if ($totalCount -gt 0) { $totalCount } else { 1 }
    $criticalWidth = [math]::Round(($criticalCount / $totalForWidth) * 100, 2)
    $warningWidth  = [math]::Round(($warningCount / $totalForWidth) * 100, 2)
    $noticeWidth   = [math]::Round(($noticeCount / $totalForWidth) * 100, 2)
    if ($totalCount -eq 0) { $noticeWidth = 100 }

    $html = $template
    $html = $html.Replace('__TITLE__', (Convert-ToHtmlSafe 'ADCS Expiring Certificate Report'))
    $html = $html.Replace('__GENERATED__', (Convert-ToHtmlSafe $generatedAt.ToString('yyyy-MM-dd HH:mm:ss')))
    $html = $html.Replace('__WINDOW__', (Convert-ToHtmlSafe $WindowDays))
    $html = $html.Replace('__TOTAL__', (Convert-ToHtmlSafe $totalCount))
    $html = $html.Replace('__CRITICAL__', (Convert-ToHtmlSafe $criticalCount))
    $html = $html.Replace('__WARNING__', (Convert-ToHtmlSafe $warningCount))
    $html = $html.Replace('__NOTICE__', (Convert-ToHtmlSafe $noticeCount))
    $html = $html.Replace('__SOONEST__', (Convert-ToHtmlSafe $(if ($soonestDays -lt 0) { 'None in selected window' } else { [string]$soonestDays })))
    $html = $html.Replace('__SOONEST_RAW__', (Convert-ToHtmlSafe $(if ($soonestDays -ge 0) { $soonestDays } else { 'N/A' })))
    $html = $html.Replace('__DONUT__', $chartGradient)
    $html = $html.Replace('__CSV_FILE__', (Convert-ToHtmlSafe ([System.IO.Path]::GetFileName($csvPath))))
    $html = $html.Replace('__CRITICAL_WIDTH__', [string]$criticalWidth)
    $html = $html.Replace('__WARNING_WIDTH__', [string]$warningWidth)
    $html = $html.Replace('__NOTICE_WIDTH__', [string]$noticeWidth)
    $html = $html.Replace('__REPORT_JSON__', $json)

    Set-Content -Path $htmlPath -Value $html -Encoding UTF8

    [Console]::WriteLine($htmlPath)

    if ($OpenReport) {
        Start-Process $htmlPath
    }
}
catch {
    $errorMessage = $_.Exception.Message
    try {
        Write-ADCSMetricLog -Message "Get-ADCSExpiringCertificateReport60Days-Styled failed: $errorMessage"
    }
    catch {
    }
    throw
}
