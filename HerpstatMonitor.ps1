<#
.SYNOPSIS
  Polls one or more Herpstat thermostats, logs key readings, sends alerts on failures
  (SMTP email and Textbelt SMS with rate limiting plus a shared consecutive failure threshold),
  emails twice-daily temperature summaries, sends summary deviation alerts when average
  probe temperature drifts from set temperature, writes a per-run verbose action log
  with retention, and pings Healthchecks.io for start/success/failure.

.DESCRIPTION
  - For each device IP in Devices:
      * Ping to verify reachability.
      * If reachable, fetch RAWSTATUS JSON and extract OutputNumber, OutputName,
        ProbeTemp, SetTemp, PowerOutput for outputs 1 and 2.
      * Write one CSV row per output (Timestamp, TimestampISO, fields, DeviceIP).
  - Alerts:
      * Email and SMS only after the same device reaches the consecutive failure threshold.
      * Recovery notifications are email-only.
      * SMS messages are sanitized to remove links/IPs/slash-dates, and are rate-limited
        per alert category to at most one text every N hours (default 8). SMS suppression is logged.
      * Email is sent through Gmail SMTP using an app password.
  - Twice-daily summary email near 8:00 AM and 8:00 PM (configurable window), combining
    data across devices.
  - Summary deviation alerts: when an output's average probe temp differs from its
    average set temp by the configured threshold (default 1.0 degree), send email/SMS,
    plus an email-only recovery when the output returns to normal.
  - Probe sanity alerts: when a live probe reading falls outside the configured sanity
    range, send email/SMS, plus an email-only recovery when the reading returns to normal.
  - CSV maintenance: keep last N days in live CSV (default 180); archive older rows to ZIP.
  - Verbose per-run .txt log under LogDir\Verbose; keeps newest 50 files.
  - Healthchecks.io integration:
      * /start at run start
      * success if all devices stay below the consecutive failure threshold
      * /fail with short message if any device reaches the threshold
  - On-demand:
      * -ForceStatusNow: send a live status email (for all devices) with SMS fallback if email fails.
      * -ForceSummaryNow: send a 12-hour summary immediately (all devices) with SMS fallback.
      * -SendTestAlertsNow: send a manual test issue alert and a manual test recovery alert, then exit.
      * -SendTestSummaryDeviationNow: send a manual summary deviation issue alert and recovery alert, then exit.
      * -SendTestProbeSanityNow: send a manual probe sanity issue alert and recovery alert, then exit.
      * -ResetAlertStates: clear saved alert/cooldown state so first-occurrence alerts can be tested again.
      * -SkipHealthchecks: skip Healthchecks.io pings during the current run.
      * -DryRunAlerts: suppress email/SMS/Healthchecks side effects and log what would happen.

.PARAMETER Devices
  Array of device IPs to poll. Default includes .50 and .51.

.PARAMETER SummaryDeviationThreshold
  Sends a summary deviation email/SMS when an output's average probe temp differs from
  its average set temp by at least this many degrees during the scheduled summary window.

.PARAMETER ProbeSanityEnabled
  Enables immediate probe sanity checks against ProbeTempMin and ProbeTempMax.

.PARAMETER ProbeTempMin
  Minimum acceptable live probe temperature before a probe sanity issue alert is raised.

.PARAMETER ProbeTempMax
  Maximum acceptable live probe temperature before a probe sanity issue alert is raised.

.PARAMETER DryRunAlerts
  Suppresses email, SMS, and Healthchecks sends and logs the attempted actions instead.

.PARAMETER SendTestAlertsNow
  Sends a manual issue alert and recovery alert immediately, then exits.

.PARAMETER SendTestSummaryDeviationNow
  Sends a manual summary deviation issue alert and recovery alert immediately, then exits.

.PARAMETER SendTestProbeSanityNow
  Sends a manual probe sanity issue alert and recovery alert immediately, then exits.

.PARAMETER ResetAlertStates
  Clears saved alert state files so issue alerts can be triggered again during testing.

.PARAMETER SkipHealthchecks
  Skips Healthchecks.io pings during the current run.

.EXAMPLE
  .\HerpstatMonitor.ps1 -SendTestAlertsNow -DryRunAlerts
  Safe dry run of the basic issue/recovery alert test path.

.EXAMPLE
  .\HerpstatMonitor.ps1 -SendTestAlertsNow
  Sends a real test issue email/SMS and a real test recovery email, then exits.

.EXAMPLE
  .\HerpstatMonitor.ps1 -SendTestSummaryDeviationNow -DryRunAlerts
  Safe dry run of the summary deviation issue/recovery test path.

.EXAMPLE
  .\HerpstatMonitor.ps1 -SendTestProbeSanityNow -DryRunAlerts
  Safe dry run of the probe sanity issue/recovery test path.

.EXAMPLE
  .\HerpstatMonitor.ps1 -ResetAlertStates -DryRunAlerts
  Shows which saved alert state files would be cleared before retesting.

.EXAMPLE
  .\HerpstatMonitor.ps1 -ForceStatusNow -Devices @() -SkipHealthchecks
  Sends a manual status email without contacting devices.

.EXAMPLE
  .\HerpstatMonitor.ps1 -ForceSummaryNow -Devices @() -SkipHealthchecks
  Sends a manual summary email using existing CSV history without contacting devices.

.NOTES
  User-editable defaults are grouped in the configuration section of the param block
  below so setup is easier for new users.

#>

param(
    # ================= USER CONFIGURATION =================
    # Devices and names
    [string[]]$Devices = @("192.168.1.50","192.168.1.51"),
    [hashtable]$DeviceNames = @{
        '192.168.1.50' = 'Herpstat1'
        '192.168.1.51' = 'Herpstat2'
    },

    # Logging and polling
    [string]$LogDir   = $(if ([string]::IsNullOrWhiteSpace([Environment]::GetFolderPath('Desktop'))) { Join-Path $env:USERPROFILE 'Desktop\Herpstat' } else { Join-Path ([Environment]::GetFolderPath('Desktop')) 'Herpstat' }),
    [string]$LogFile  = "herpstat_log.csv",
    [int]$PingCount   = 1,
    [ValidateRange(1, 2147483647)][int]$FailureThreshold = 2,
    [int]$RetentionDays = 180,
    [string]$ArchiveDir = "$LogDir\Archive",
    [string]$VerboseDir = "$LogDir\Verbose",
    [int]$VerboseKeep = 50,

    # Summary and alert thresholds
    [int]$SummaryHourAM = 8,
    [ValidateRange(0, 59)][int]$SummaryMinuteAM = 0,
    [int]$SummaryHourPM = 20,
    [ValidateRange(0, 59)][int]$SummaryMinutePM = 0,
    [int]$SummaryWindowMinutes = 20,
    [ValidateRange(0.1, 2147483647)][double]$SummaryDeviationThreshold = 1.0,
    [bool]$ProbeSanityEnabled = $true,
    [double]$ProbeTempMin = 20.0,
    [double]$ProbeTempMax = 130.0,

    # Textbelt SMS
    [string]$SmsTo = "",
    [string]$TextbeltApiKey = "",
    [string]$TextbeltEndpoint = "https://textbelt.com/text",
    [bool]$SmsRateLimitEnabled = $true,
    [int]$SmsRateLimitHours = 8,

    # Healthchecks.io
    [string]$HealthchecksUrl = "",

    # Email via Gmail SMTP
    [string]$MailFrom = "youraddress@gmail.com",
    [string]$MailTo = "youraddress@gmail.com",
    [string]$SmtpServer = "smtp.gmail.com",
    [int]$SmtpPort = 587,
    [bool]$UseSsl = $true,
    [string]$SmtpUsername = "youraddress@gmail.com",
    [string]$SmtpAppPassword = "",

    # ================= RUNTIME SWITCHES =================
    [switch]$ForceStatusNow,
    [switch]$ForceSummaryNow,
    [switch]$DryRunAlerts,
    [switch]$SendTestAlertsNow,
    [switch]$SendTestSummaryDeviationNow,
    [switch]$SendTestProbeSanityNow,
    [switch]$ResetAlertStates,
    [switch]$SkipHealthchecks
)

function Get-DeviceFriendlyName {
    param([Parameter(Mandatory)][string]$Ip)
    if ($DeviceNames.ContainsKey($Ip)) { return $DeviceNames[$Ip] }
    return $Ip   # fallback if an IP isn't mapped
}

# ================= Paths =================
$null = New-Item -ItemType Directory -Path $LogDir -Force -ErrorAction SilentlyContinue
$null = New-Item -ItemType Directory -Path $ArchiveDir -Force -ErrorAction SilentlyContinue
$null = New-Item -ItemType Directory -Path $VerboseDir -Force -ErrorAction SilentlyContinue

$script:logPath        = Join-Path $LogDir $LogFile
$script:statePath      = Join-Path $LogDir "last_summary.json"
$script:smsStatePath   = Join-Path $LogDir "last_sms.json"
$script:failStatePath  = Join-Path $LogDir "failures.json"
$script:summaryAlertStatePath = Join-Path $LogDir "summary_deviation_alerts.json"
$script:probeAlertStatePath   = Join-Path $LogDir "probe_sanity_alerts.json"
$script:verbosePath    = $null  # set by Initialize-VerboseLog
$script:smtpCredential = $null

if (-not [string]::IsNullOrWhiteSpace($SmtpUsername) -and -not [string]::IsNullOrWhiteSpace($SmtpAppPassword)) {
    $script:smtpCredential = New-Object System.Management.Automation.PSCredential(
        $SmtpUsername,
        (ConvertTo-SecureString $SmtpAppPassword -AsPlainText -Force)
    )
}

# ================= Verbose logging helpers =================
function Initialize-VerboseLog {
    [CmdletBinding()]
    param()
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss_fff'
    $path  = Join-Path $VerboseDir ("herpstat_run_{0}.txt" -f $stamp)
    $script:verbosePath = $path
    $hcLine = if ($SkipHealthchecks) { "Healthchecks: Skipped=True Url=$HealthchecksUrl" } elseif ([string]::IsNullOrWhiteSpace($HealthchecksUrl)) { "Healthchecks: Url=" } else { "Healthchecks: Url=$HealthchecksUrl" }
    $header = @(
        "===== Herpstat run started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') =====",
        "Devices: " + ($Devices -join ', '),
        "LogDir: $LogDir",
        "CSV: $script:logPath",
        "ArchiveDir: $ArchiveDir",
        "VerboseDir: $VerboseDir",
        "Summary targets: {0}:{1:d2} and {2}:{3:d2}, window: +/- {4} min" -f $SummaryHourAM, $SummaryMinuteAM, $SummaryHourPM, $SummaryMinutePM, $SummaryWindowMinutes,
        "Summary deviation alert threshold: $SummaryDeviationThreshold degree(s)",
        "Probe sanity: Enabled=$ProbeSanityEnabled, Min=$ProbeTempMin, Max=$ProbeTempMax",
        "RetentionDays: $RetentionDays, VerboseKeep: $VerboseKeep",
        "SMS rate limiting: Enabled=$([bool]$SmsRateLimitEnabled), WindowHours=$SmsRateLimitHours, PerCategory=True",
        "DryRunAlerts: $DryRunAlerts",
        "ResetAlertStates: $ResetAlertStates, SkipHealthchecks: $SkipHealthchecks",
        $hcLine,
        "============================================================"
    ) -join [Environment]::NewLine
    $null = $header | Out-File -FilePath $path -Encoding UTF8
}
function Write-ActionLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR')][string]$Level = 'INFO',
        [switch]$AlsoConsole
    )
    if (-not $script:verbosePath) { return }
    $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line  = "{0} [{1}] {2}" -f $stamp, $Level, $Message
    Add-Content -Path $script:verbosePath -Value $line
    if ($AlsoConsole) {
        switch ($Level) {
            'INFO'  { Write-Host $Message }
            'WARN'  { Write-Warning $Message }
            'ERROR' { Write-Error $Message }
        }
    }
}
function Invoke-VerboseLogRetention {
    [CmdletBinding()]
    param()
    try {
        $files = Get-ChildItem -Path $VerboseDir -File -Filter 'herpstat_run_*.txt' | Sort-Object LastWriteTime -Descending
        if ($files.Count -gt $VerboseKeep) {
            $toRemove = $files | Select-Object -Skip $VerboseKeep
            foreach ($f in $toRemove) { Remove-Item -Path $f.FullName -Force -ErrorAction SilentlyContinue }
            Write-ActionLog -Message ("Verbose retention removed {0} old file(s)" -f $toRemove.Count)
        } else {
            Write-ActionLog -Message "Verbose retention: nothing to remove"
        }
    } catch {
        Write-ActionLog -Message ("Verbose retention error: {0}" -f $_.Exception.Message) -Level ERROR
    }
}
function Get-OutputAlertKey {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DeviceIp,
        [Parameter(Mandatory)][int]$OutputNumber,
        [Parameter(Mandatory)][string]$OutputName
    )
    return "{0}|{1}|{2}" -f $DeviceIp, $OutputNumber, $OutputName
}
function Get-BooleanAlertState {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path $Path)) { return @{} }

    try {
        $obj = Get-Content -Path $Path -Raw | ConvertFrom-Json
        if ($null -eq $obj) { return @{} }

        $state = @{}
        foreach ($prop in $obj.PSObject.Properties) {
            $state[$prop.Name] = [bool]$prop.Value
        }
        return $state
    } catch {
        Write-ActionLog -Message ("Alert state read failed for {0}: {1}" -f $Path, $_.Exception.Message) -Level WARN
        return @{}
    }
}
function Set-BooleanAlertState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][hashtable]$State
    )

    if ($DryRunAlerts) {
        Write-ActionLog -Message ("DRY RUN state skipped for {0}" -f $Path)
        return
    }

    $obj = [pscustomobject]@{}
    foreach ($key in ($State.Keys | Sort-Object)) {
        Add-Member -InputObject $obj -NotePropertyName $key -NotePropertyValue ([bool]$State[$key]) -Force
    }
    $obj | ConvertTo-Json -Compress | Set-Content -Path $Path -Encoding UTF8
}
function Reset-AlertStates {
    [CmdletBinding()]
    param()

    $paths = @(
        $script:failStatePath,
        $script:smsStatePath,
        $script:summaryAlertStatePath,
        $script:probeAlertStatePath
    )

    foreach ($path in $paths) {
        if (Test-Path $path) {
            if ($DryRunAlerts) {
                Write-ActionLog -Message ("DRY RUN alert state reset skipped for {0}" -f $path)
            } else {
                Remove-Item -Path $path -Force -ErrorAction SilentlyContinue
                Write-ActionLog -Message ("Alert state cleared: {0}" -f $path)
            }
        } else {
            Write-ActionLog -Message ("Alert state not present: {0}" -f $path)
        }
    }
}

# ================= Healthchecks.io helpers =================
function Invoke-HealthcheckPing {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$BaseUrl,
        [ValidateSet('start','success','fail')][string]$Kind = 'success',
        [string]$BodyText = $null
    )

    if ([string]::IsNullOrWhiteSpace($BaseUrl)) { return }
    if ($SkipHealthchecks) {
        Write-ActionLog -Message ("Healthchecks ping skipped by switch: {0}" -f $Kind)
        return
    }
    if ($DryRunAlerts) {
        Write-ActionLog -Message ("DRY RUN Healthchecks ping skipped: {0}" -f $Kind)
        return
    }

    $suffix = switch ($Kind) {
        'start' { '/start' }
        'fail'  { '/fail' }
        default { '' }
    }
    $url = ($BaseUrl.TrimEnd('/')) + $suffix

    # -----------------------------------------
    # DIAGNOSTICS BEFORE ATTEMPTS
    # -----------------------------------------
    try {
        $dns = Resolve-DnsName -Name 'hc-ping.com' -ErrorAction Stop
        Write-ActionLog -Message ("DNS for hc-ping.com: " + ($dns.IPAddress -join ', '))
    }
    catch {
        Write-ActionLog -Message ("DNS lookup failed for hc-ping.com: {0}" -f $_.Exception.Message) -Level WARN
    }

    try {
        $tn = Test-NetConnection -ComputerName 'hc-ping.com' -Port 443 -InformationLevel Detailed -WarningAction SilentlyContinue
        Write-ActionLog -Message ("Test-NetConnection hc-ping.com: TcpTestSucceeded={0}" -f $tn.TcpTestSucceeded)
    }
    catch {
        Write-ActionLog -Message ("Test-NetConnection failed for hc-ping.com: {0}" -f $_.Exception.Message) -Level WARN
    }

    # -----------------------------------------
    # RETRY LOOP (3 attempts)
    # -----------------------------------------
    $maxAttempts = 3
    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        try {
            Write-ActionLog -Message ("Healthchecks ping attempt {0}/{1} ({2})" -f $attempt, $maxAttempts, $Kind)

            if ($BodyText) {
                Invoke-RestMethod -Uri $url -Method Post -Body $BodyText -ContentType 'text/plain' -TimeoutSec 10 | Out-Null
            } else {
                Invoke-RestMethod -Uri $url -Method Get -TimeoutSec 10 | Out-Null
            }

            Write-ActionLog -Message ("Healthchecks ping ok: {0}" -f $Kind)
            return
        }
        catch {
            $ex = $_.Exception
            $inner = if ($ex.InnerException) { $ex.InnerException.Message } else { "<none>" }

            Write-ActionLog -Message (
                "Healthchecks ping failed ({0}) attempt {1}: {2} | Inner: {3}" -f $Kind, $attempt, $ex.Message, $inner
            ) -Level WARN -AlsoConsole

            # Wait before retrying
            Start-Sleep -Seconds 5
        }
    }

    # -----------------------------------------
    # After all attempts fail — secondary check
    # -----------------------------------------
    Write-ActionLog -Message ("Healthchecks ping ultimately failed after {0} attempts ({1})" -f $maxAttempts, $Kind) -Level ERROR -AlsoConsole

    try {
        $googleTest = Invoke-RestMethod -Uri 'https://www.google.com/generate_204' -TimeoutSec 5 -ErrorAction Stop
        Write-ActionLog -Message "Secondary connectivity check (Google 204) succeeded"
    }
    catch {
        Write-ActionLog -Message ("Secondary connectivity check failed: {0}" -f $_.Exception.Message) -Level WARN
    }
}

# ================= SMS timestamp and sanitizer =================
function Get-SmsSafeTimestamp {
    [CmdletBinding()]
    param([datetime]$When = (Get-Date))
    return $When.ToString('MM-dd-yyyy hhmm tt')  # example 09-21-2025 0952 PM
}
function ConvertTo-SmsSafeText {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Text)
    $t = $Text
    $t = [regex]::Replace($t, '(?i)\b((?:https?|ftp)://\S+|www\.\S+)\b', '[link removed]')
    $t = [regex]::Replace($t, '(?i)\b[a-z0-9.-]+\.(com|org|net|io|gov|edu|co|us|uk|ca|info|biz|dev|tech|me|ly|ai)\b', '[link removed]')
    $t = [regex]::Replace($t, '\b\d{1,3}(?:\.\d{1,3}){3}(?:/\S+)?\b', '[ip removed]')
    $t = [regex]::Replace($t, '\b(?:[A-Fa-f0-9]{1,4}:){2,7}[A-Fa-f0-9]{1,4}\b', '[ip removed]')
    $t = [regex]::Replace($t, '\b\d{1,2}/\d{1,2}/\d{2,4}(?:\s+\d{1,2}:\d{2}(?::\d{2})?\s*(?:AM|PM)?)?\b', '[date]')
    $t = [regex]::Replace($t, '\s+', ' ')
    return $t.Trim()
}

# ================= Email helpers =================
function Get-ExceptionDetailText {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$ErrorRecord)

    $parts = @()
    $ex = $ErrorRecord.Exception
    if ($ex -and $ex.Message) {
        $parts += [string]$ex.Message
    }

    $response = $null
    if ($ex -and $ex.PSObject.Properties.Name -contains 'Response') {
        $response = $ex.Response
    }

    if ($response) {
        try {
            if ($response.PSObject.Properties.Name -contains 'StatusCode' -and $null -ne $response.StatusCode) {
                $parts += ("HTTP {0}" -f [int]$response.StatusCode)
            }
            if ($response.PSObject.Properties.Name -contains 'StatusDescription' -and $response.StatusDescription) {
                $parts += [string]$response.StatusDescription
            }

            $stream = $response.GetResponseStream()
            if ($stream) {
                $reader = New-Object System.IO.StreamReader($stream)
                try {
                    $responseText = $reader.ReadToEnd()
                    if (-not [string]::IsNullOrWhiteSpace($responseText)) {
                        $compact = ($responseText -replace '\s+', ' ').Trim()
                        if ($compact.Length -gt 500) { $compact = $compact.Substring(0, 500) + '...' }
                        $parts += ("Response={0}" -f $compact)
                    }
                } finally {
                    $reader.Dispose()
                    $stream.Dispose()
                }
            }
        } catch {}
    }

    if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
        $detail = (($ErrorRecord.ErrorDetails.Message -replace '\s+', ' ').Trim())
        if ($detail) { $parts += ("ErrorDetails={0}" -f $detail) }
    }

    $parts = @($parts | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Select-Object -Unique)
    if ($parts.Count -gt 0) { return ($parts -join ' | ') }
    return "Unknown email send failure."
}
function Send-EmailRaw {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Subject,
        [Parameter(Mandatory)]$Body,
        [switch]$AsHtml
    )
    if ($null -eq $Body) { $Body = "" }
    elseif ($Body -is [System.Array]) { $Body = ($Body -join [Environment]::NewLine) }
    else { $Body = [string]$Body }

    if ($DryRunAlerts) {
        Write-ActionLog -Message ("DRY RUN email skipped: {0}" -f $Subject) -Level WARN
        return $true
    }

    try {
        if (-not $script:smtpCredential) {
            throw "SMTP credentials are missing."
        }
        $params = @{
            To          = $MailTo
            From        = $MailFrom
            Subject     = $Subject
            Body        = $Body
            SmtpServer  = $SmtpServer
            Port        = $SmtpPort
            UseSsl      = $UseSsl
            Credential  = $script:smtpCredential
            ErrorAction = 'Stop'
        }
        if ($AsHtml) { $params['BodyAsHtml'] = $true }

        Send-MailMessage @params
        Write-ActionLog -Message ("Email sent via SMTP: {0}" -f $Subject)
        return $true
    } catch {
        $detail = Get-ExceptionDetailText -ErrorRecord $_
        Write-ActionLog -Message ("Email send failed: {0}" -f $detail) -Level ERROR
        return $false
    }
}
function Send-EmailWithSmsFallback {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Subject,
        [Parameter(Mandatory)]$Body,
        [switch]$AsHtml,
        [string]$SmsMessage
    )
    $ok = Send-EmailRaw -Subject $Subject -Body $Body -AsHtml:$AsHtml
    if (-not $ok) {
        $when = Get-SmsSafeTimestamp
        if ([string]::IsNullOrWhiteSpace($SmsMessage)) { $SmsMessage = "Email failed at $when. Subject: $Subject" }
        $null = Send-TextbeltAlert -Message $SmsMessage -FailureContext "Email failed, SMS fallback" -AlertCategory 'email_fallback'
        Write-ActionLog -Message ("SMS fallback evaluated for failed email: {0}" -f $Subject) -Level WARN -AlsoConsole
    }
    return $ok
}
function Send-EmailAlert {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Subject,[Parameter(Mandatory)][string]$Body)
    $null = Send-EmailRaw -Subject $Subject -Body $Body
}

# ================= SMS state and sender =================
function Get-LastSmsInfo {
    param([string]$AlertCategory = 'general')

    $state = @{}
    if (Test-Path $script:smsStatePath) {
        try {
            $obj = Get-Content -Path $script:smsStatePath -Raw | ConvertFrom-Json
            if ($obj) {
                if ($obj.PSObject.Properties.Name -contains 'Categories' -and $obj.Categories) {
                    foreach ($prop in $obj.Categories.PSObject.Properties) {
                        $state[$prop.Name] = [string]$prop.Value
                    }
                }
                if ($obj.PSObject.Properties.Name -contains 'LastSmsISO' -and $obj.LastSmsISO -and -not $state.ContainsKey('legacy')) {
                    $state['legacy'] = [string]$obj.LastSmsISO
                }
            }
        } catch {
            Write-ActionLog -Message ("SMS state read failed: {0}" -f $_.Exception.Message) -Level WARN
        }
    }

    $iso = $null
    if ($state.ContainsKey($AlertCategory)) { $iso = $state[$AlertCategory] }
    elseif ($state.ContainsKey('legacy')) { $iso = $state['legacy'] }

    if ($iso) {
        return [pscustomobject]@{
            AlertCategory = $AlertCategory
            LastSmsISO    = $iso
        }
    }
    return $null
}
function Set-LastSmsInfo {
    param(
        [datetime]$LastSent,
        [string]$AlertCategory = 'general'
    )

    if ($DryRunAlerts) {
        Write-ActionLog -Message ("DRY RUN SMS state skipped for category '{0}'" -f $AlertCategory)
        return
    }

    $state = @{}
    if (Test-Path $script:smsStatePath) {
        try {
            $obj = Get-Content -Path $script:smsStatePath -Raw | ConvertFrom-Json
            if ($obj -and $obj.PSObject.Properties.Name -contains 'Categories' -and $obj.Categories) {
                foreach ($prop in $obj.Categories.PSObject.Properties) {
                    $state[$prop.Name] = [string]$prop.Value
                }
            }
        } catch {
            Write-ActionLog -Message ("SMS state merge failed: {0}" -f $_.Exception.Message) -Level WARN
        }
    }

    $state[$AlertCategory] = $LastSent.ToString("o")
    if ($state.ContainsKey('legacy')) { $null = $state.Remove('legacy') }

    $cats = [pscustomobject]@{}
    foreach ($key in ($state.Keys | Sort-Object)) {
        Add-Member -InputObject $cats -NotePropertyName $key -NotePropertyValue ([string]$state[$key]) -Force
    }
    $obj = [pscustomobject]@{ Categories = $cats }
    $obj | ConvertTo-Json -Depth 4 | Set-Content -Path $script:smsStatePath -Encoding UTF8
    Write-ActionLog -Message ("SMS state saved: Category={0} LastSms={1}" -f $AlertCategory, $LastSent)
}
function Send-TextbeltAlert {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Message,
        [switch]$EmailOnFailure = $true,
        [string]$FailureContext = "",
        [string]$AlertCategory = 'general'
    )
    if ([string]::IsNullOrWhiteSpace($SmsTo) -or [string]::IsNullOrWhiteSpace($TextbeltApiKey)) {
        Write-ActionLog -Message "SMS disabled (SmsTo or TextbeltApiKey not set)"
        return $false
    }

    $clean = ConvertTo-SmsSafeText -Text $Message
    if ($DryRunAlerts) {
        Write-ActionLog -Message ("DRY RUN SMS skipped [{0}]: {1}" -f $AlertCategory, $clean) -Level WARN
        return $true
    }

    # Rate limiting: suppress if within window
    if ($SmsRateLimitEnabled) {
        $state = Get-LastSmsInfo -AlertCategory $AlertCategory
        if ($state -and $state.LastSmsISO) {
            try {
                $last = [datetime]::Parse([string]$state.LastSmsISO)
                $elapsed = (Get-Date) - $last
                if ($elapsed.TotalHours -lt [double]$SmsRateLimitHours) {
                    $remaining = [math]::Ceiling(([timespan]::FromHours($SmsRateLimitHours) - $elapsed).TotalMinutes)
                    Write-ActionLog -Message ("SMS suppressed by rate limit for category '{0}'; {1} minute(s) remaining" -f $AlertCategory, $remaining) -AlsoConsole
                    return $true
                }
            } catch {
                Write-ActionLog -Message ("SMS state parse error: {0}" -f $_.Exception.Message) -Level WARN
            }
        }
    }
    try {
        $body = @{ phone = $SmsTo; message = $clean; key = $TextbeltApiKey }
        $resp = Invoke-RestMethod -Uri $TextbeltEndpoint -Method Post -Body $body -ContentType 'application/x-www-form-urlencoded' -TimeoutSec 10
        if ($resp -and $resp.success -eq $true) {
            Write-ActionLog -Message ("SMS sent [{0}]" -f $AlertCategory) -AlsoConsole
            Set-LastSmsInfo -LastSent (Get-Date) -AlertCategory $AlertCategory
            return $true
        } else {
            $detail = if ($resp) { $resp | ConvertTo-Json -Compress } else { "no response" }
            Write-ActionLog -Message ("SMS failed [{0}]: {1}" -f $AlertCategory, $detail) -Level WARN -AlsoConsole
            if ($EmailOnFailure) {
                $subject = "NOTICE: SMS send failed"
                $when    = Get-Date -Format 'MM/dd/yyyy hh:mm:ss tt'
                $ctxLine = if ($FailureContext) { "Context: $FailureContext`r`n" } else { "" }
                $bodyTxt = "An SMS failed to send at $when.`r`n${ctxLine}Message attempted: $clean`r`nResponse: $detail"
                $null = Send-EmailRaw -Subject $subject -Body $bodyTxt
            }
            return $false
        }
    } catch {
        $err = $_.Exception.Message
        Write-ActionLog -Message ("SMS exception [{0}]: {1}" -f $AlertCategory, $err) -Level ERROR -AlsoConsole
        if ($EmailOnFailure) {
            $subject = "NOTICE: SMS send failed (exception)"
            $when    = Get-Date -Format 'MM/dd/yyyy hh:mm:ss tt'
            $ctxLine = if ($FailureContext) { "Context: $FailureContext`r`n" } else { "" }
            $bodyTxt = "An SMS failed to send at $when due to an exception.`r`n${ctxLine}Message attempted: $clean`r`nError: $err"
            $null = Send-EmailRaw -Subject $subject -Body $bodyTxt
        }
        return $false
    }
}

# Exclude any outputs whose name is "Nickname", "Nickname2", etc. (case-insensitive)
function Test-ExcludeName {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return $false }
    return ($Name -match '^(?i)nickname(\d+)?$')
}

# === Consecutive failure tracking (per device) ===
function Get-FailureState {
    # Returns a [hashtable] mapping IP -> int failureCount
    if (Test-Path $script:failStatePath) {
        try {
            $obj = Get-Content -Path $script:failStatePath -Raw | ConvertFrom-Json
            if ($null -eq $obj) { return @{} }

            # Convert PSCustomObject to hashtable
            $ht = @{}
            foreach ($prop in $obj.PSObject.Properties) {
                # Coerce to [int] just in case
                $ht[$prop.Name] = [int]$prop.Value
            }
            return $ht
        } catch {
            # If file is corrupt or unreadable, start fresh
            return @{}
        }
    }
    return @{}
}
function Set-FailureState {
    param([hashtable]$State)
    # Serialize hashtable back to a simple JSON object
    $psobj = [pscustomobject]@{}
    foreach ($k in $State.Keys) {
        Add-Member -InputObject $psobj -NotePropertyName $k -NotePropertyValue ([int]$State[$k]) -Force
    }
    $psobj | ConvertTo-Json -Compress | Set-Content -Path $script:failStatePath -Encoding UTF8
}
function Register-DeviceResult {
    <#
      Updates consecutive failure counts for a device and triggers email/SMS
      only after failures reach the configured threshold. Returns an object
      with FailureCount and ThresholdReached.
    #>
    param(
        [string]$DeviceIp,
        [bool]$Success,
        [string]$SmsMessage,
        [string]$FailureContext,
        [string]$EmailSubject,
        [string]$EmailBody
    )

    $state = Get-FailureState
    $previousCount = if ($state.ContainsKey($DeviceIp)) { [int]$state[$DeviceIp] } else { 0 }

    if ($Success) {
        if ($previousCount -ge $FailureThreshold) {
            $deviceName = Get-DeviceFriendlyName -Ip $DeviceIp
            $now12      = Get-Date -Format 'MM/dd/yyyy hh:mm:ss tt'
            $subject    = "RECOVERY: Herpstat responding again ($deviceName / $DeviceIp)"
            $body       = "Herpstat $deviceName ($DeviceIp) responded normally again at $now12 after $previousCount consecutive failure(s)."
            Send-EmailAlert -Subject $subject -Body $body
            Write-ActionLog -Message ("Device recovery email sent for {0} after {1} consecutive failure(s)" -f $DeviceIp, $previousCount)
        }

        # Reset counter on success
        $state[$DeviceIp] = 0
        Set-FailureState -State $state
        Write-ActionLog -Message ("Failure counter reset for {0}" -f $DeviceIp)
        return [pscustomobject]@{
            FailureCount     = 0
            ThresholdReached = $false
        }
    }

    # Increment on failure
    if (-not $state.ContainsKey($DeviceIp)) { $state[$DeviceIp] = 0 }
    $state[$DeviceIp] = [int]$state[$DeviceIp] + 1
    $count = [int]$state[$DeviceIp]
    Set-FailureState -State $state

    $thresholdReached = ($count -ge $FailureThreshold)
    Write-ActionLog -Message ("Failure counter for {0}: {1}/{2}" -f $DeviceIp, $count, $FailureThreshold) -Level WARN
    if ($thresholdReached) {
        if (-not [string]::IsNullOrWhiteSpace($EmailSubject) -and -not [string]::IsNullOrWhiteSpace($EmailBody)) {
            Send-EmailAlert -Subject $EmailSubject -Body $EmailBody
        }
        $null = Send-TextbeltAlert -Message $SmsMessage -FailureContext $FailureContext -AlertCategory 'device_issue'
    } else {
        Write-ActionLog -Message ("Notifications held for {0} until {1} consecutive failures" -f $DeviceIp, $FailureThreshold) -Level WARN
    }
    return [pscustomobject]@{
        FailureCount     = $count
        ThresholdReached = $thresholdReached
    }
}

# ================= CSV logging and parsing =================
function Write-LogRow {
    [CmdletBinding()]
    param(
        [int]$OutputNumber,[string]$Nickname,[double]$ProbeTemp,[double]$SetTemp,[int]$PowerOutput,[string]$DeviceIp
    )
    $row = [pscustomobject]@{
        Timestamp    = (Get-Date -Format "MM/dd/yyyy hh:mm:ss tt")
        TimestampISO = (Get-Date).ToString("s")
        OutputNumber = $OutputNumber
        OutputName   = $Nickname
        ProbeTemp    = $ProbeTemp
        SetTemp      = $SetTemp
        PowerOutput  = $PowerOutput
        DeviceIP     = $DeviceIp
    }
    if (-not (Test-Path $script:logPath)) {
        $row | Export-Csv -Path $script:logPath -NoTypeInformation -Encoding UTF8
    } else {
        $row | Export-Csv -Path $script:logPath -NoTypeInformation -Encoding UTF8 -Append
    }
    Write-ActionLog -Message ("Logged CSV row ip={0} o{1} '{2}' probe={3} set={4} power={5}%" -f $DeviceIp,$OutputNumber,$Nickname,$ProbeTemp,$SetTemp,$PowerOutput)
}
function Get-HerpstatOutput {
    [CmdletBinding()]
    param([object]$obj, [int]$n)
    if (-not $obj) { return $null }
    [pscustomobject]@{
        OutputNumber = $n
        OutputName   = $obj.outputnickname
        ProbeTemp    = [double]$obj.probereadingTEMP
        SetTemp      = [double]$obj.currentsetting
        PowerOutput  = [int]$obj.poweroutput
    }
}

# ================= Summary state and schedule =================
function Get-LastSummaryInfo {
    if (Test-Path $script:statePath) {
        try { return (Get-Content -Path $script:statePath -Raw | ConvertFrom-Json) } catch {}
    }
    return $null
}
function Set-LastSummaryInfo {
    param([datetime]$LastSent, [datetime]$LastTarget)
    if ($DryRunAlerts) {
        Write-ActionLog -Message ("DRY RUN summary state skipped: LastSent={0} LastTarget={1}" -f $LastSent, $LastTarget)
        return
    }
    $obj = [pscustomobject]@{ LastSentISO = $LastSent.ToString("o"); LastTargetISO = $LastTarget.ToString("o") }
    $obj | ConvertTo-Json | Set-Content -Path $script:statePath -Encoding UTF8
    Write-ActionLog -Message ("Summary state saved: LastSent={0} LastTarget={1}" -f $LastSent, $LastTarget)
}
function Get-CurrentTargetTime {
    param(
        [datetime]$Now,
        [int]$HourAM,
        [int]$HourPM,
        [int]$MinuteAM = 0,
        [int]$MinutePM = 0
    )
    $today = $Now.Date
    $targetAM = $today.AddHours($HourAM).AddMinutes($MinuteAM)
    $targetPM = $today.AddHours($HourPM).AddMinutes($MinutePM)
    if     ($Now -ge $targetPM) { return $targetPM }
    elseif ($Now -ge $targetAM) { return $targetAM }
    else { return $today.AddDays(-1).AddHours($HourPM).AddMinutes($MinutePM) }
}
function Test-SummaryWindowEligibility {
    [CmdletBinding()]
    param([datetime]$Now,[int]$WindowMinutes,[datetime]$Target,[object]$State)
    $minutesFromTarget = [math]::Abs(($Now - $Target).TotalMinutes)
    if ($minutesFromTarget -gt $WindowMinutes) { return $false }
    if ($State -and $State.LastTargetISO) {
        try { $lastTarget = [datetime]::Parse([string]$State.LastTargetISO); if ([datetime]::Compare($lastTarget,$Target) -eq 0) { return $false } } catch {}
    }
    return $true
}
function Get-RecentLogRows {
    param([datetime]$Since,[datetime]$Until)
    if (-not (Test-Path $script:logPath)) { return @() }
    $rows = Import-Csv -Path $script:logPath
    foreach ($r in $rows) {
        $dt = $null
        if ($r.PSObject.Properties.Name -contains 'TimestampISO' -and $r.TimestampISO) { try { $dt = [datetime]::Parse([string]$r.TimestampISO) } catch {} }
        if (-not $dt -and $r.Timestamp) { try { $dt = [datetime]::ParseExact([string]$r.Timestamp,'MM/dd/yyyy hh:mm:ss tt',$null) } catch {} }
        Add-Member -InputObject $r -NotePropertyName _dt -NotePropertyValue $dt -Force
    }
    $rows | Where-Object { $_._dt -and $_._dt -ge $Since -and $_._dt -le $Until }
}
function Get-SummaryOutputStats {
    [CmdletBinding()]
    param([object[]]$Rows)

    if ($Rows) {
        $Rows = @($Rows | Where-Object { $_ -and -not (Test-ExcludeName -Name $_.OutputName) })
    }

    if (-not $Rows -or $Rows.Count -eq 0) {
        return @()
    }

    $stats = foreach ($group in ($Rows | Group-Object { '{0}|{1}|{2}' -f $_.DeviceIP, $_.OutputNumber, $_.OutputName })) {
        $first = $group.Group | Select-Object -First 1

        $probeVals = @()
        $setVals   = @()
        foreach ($r in $group.Group) {
            try { $probeVals += [double]$r.ProbeTemp } catch {}
            try { $setVals   += [double]$r.SetTemp   } catch {}
        }

        $avgProbeRaw = if ($probeVals.Count -gt 0) { [double](($probeVals | Measure-Object -Average).Average) } else { $null }
        $avgSetRaw   = if ($setVals.Count   -gt 0) { [double](($setVals   | Measure-Object -Average).Average) } else { $null }
        $deviationRaw = if ($avgProbeRaw -ne $null -and $avgSetRaw -ne $null) { [math]::Abs($avgProbeRaw - $avgSetRaw) } else { $null }

        [pscustomobject]@{
            DeviceIP         = [string]$first.DeviceIP
            DeviceName       = Get-DeviceFriendlyName -Ip ([string]$first.DeviceIP)
            OutputNumber     = [int]$first.OutputNumber
            OutputName       = [string]$first.OutputName
            ReadingCount     = [math]::Max($probeVals.Count, $setVals.Count)
            AverageProbeTemp = if ($avgProbeRaw -ne $null) { [math]::Round($avgProbeRaw, 1) } else { $null }
            AverageSetTemp   = if ($avgSetRaw   -ne $null) { [math]::Round($avgSetRaw,   1) } else { $null }
            Deviation        = if ($deviationRaw -ne $null) { [math]::Round($deviationRaw, 1) } else { $null }
            DeviationRaw     = $deviationRaw
        }
    }

    return @($stats | Sort-Object DeviceName, OutputNumber, OutputName)
}
function Get-ProbeSanityFindings {
    [CmdletBinding()]
    param([object[]]$Rows)

    if (-not $ProbeSanityEnabled) { return @() }

    if ($Rows) {
        $Rows = @($Rows | Where-Object { $_ -and -not (Test-ExcludeName -Name $_.OutputName) })
    }

    if (-not $Rows -or $Rows.Count -eq 0) {
        return @()
    }

    $findings = foreach ($row in $Rows) {
        $issues = @()
        $probeTemp = $null
        $setTemp = $null

        try { $probeTemp = [double]$row.ProbeTemp } catch {}
        try { $setTemp = [double]$row.SetTemp } catch {}

        if ($probeTemp -eq $null) {
            $issues += "Probe temp is missing or not numeric."
        } elseif ($probeTemp -lt $ProbeTempMin -or $probeTemp -gt $ProbeTempMax) {
            $issues += ("Probe temp {0:N1} is outside configured sanity range {1:N1}-{2:N1}." -f $probeTemp, $ProbeTempMin, $ProbeTempMax)
        }

        if ($issues.Count -gt 0) {
            $deviceIp = [string]$row.DeviceIP
            $outputNumber = [int]$row.OutputNumber
            $outputName = [string]$row.OutputName
            $key = Get-OutputAlertKey -DeviceIp $deviceIp -OutputNumber $outputNumber -OutputName $outputName

            [pscustomobject]@{
                Key          = $key
                DeviceIP     = $deviceIp
                DeviceName   = Get-DeviceFriendlyName -Ip $deviceIp
                OutputNumber = $outputNumber
                OutputName   = $outputName
                ProbeTemp    = if ($probeTemp -ne $null) { [math]::Round($probeTemp, 1) } else { $null }
                SetTemp      = if ($setTemp   -ne $null) { [math]::Round($setTemp,   1) } else { $null }
                Issue        = ($issues -join ' ')
            }
        }
    }

    return @($findings)
}
function Send-ProbeSanityAlerts {
    [CmdletBinding()]
    param([object[]]$Rows)

    if (-not $ProbeSanityEnabled) {
        Write-ActionLog -Message "Probe sanity check disabled"
        return
    }

    if ($Rows) {
        $Rows = @($Rows | Where-Object { $_ -and -not (Test-ExcludeName -Name $_.OutputName) })
    }

    if (-not $Rows -or $Rows.Count -eq 0) {
        Write-ActionLog -Message "Probe sanity check skipped (no live rows)"
        return
    }

    $previousState = Get-BooleanAlertState -Path $script:probeAlertStatePath
    $currentSeenKeys = @{}
    $currentByKey = @{}

    foreach ($row in $Rows) {
        $deviceIp = [string]$row.DeviceIP
        $outputNumber = [int]$row.OutputNumber
        $outputName = [string]$row.OutputName
        $key = Get-OutputAlertKey -DeviceIp $deviceIp -OutputNumber $outputNumber -OutputName $outputName

        $currentSeenKeys[$key] = $true
        $currentByKey[$key] = [pscustomobject]@{
            Key          = $key
            DeviceIP     = $deviceIp
            DeviceName   = Get-DeviceFriendlyName -Ip $deviceIp
            OutputNumber = $outputNumber
            OutputName   = $outputName
            ProbeTemp    = [double]$row.ProbeTemp
            SetTemp      = [double]$row.SetTemp
        }
    }

    $findings = Get-ProbeSanityFindings -Rows $Rows
    $currentIssueKeys = @{}
    foreach ($finding in $findings) { $currentIssueKeys[$finding.Key] = $true }

    $newIssues = @($findings | Where-Object { -not $previousState.ContainsKey($_.Key) })
    $recoveries = @()
    foreach ($key in @($previousState.Keys)) {
        if ($currentSeenKeys.ContainsKey($key) -and -not $currentIssueKeys.ContainsKey($key)) {
            $recoveries += $currentByKey[$key]
        }
    }

    if ($newIssues.Count -gt 0) {
        $subject = if ($newIssues.Count -eq 1) {
            "ALERT: Probe sanity issue for $($newIssues[0].OutputName)"
        } else {
            "ALERT: Probe sanity issues detected ($($newIssues.Count) outputs)"
        }

        $bodyLines = @(
            "One or more live probe readings fell outside the configured sanity range of $([string]::Format('{0:N1}', $ProbeTempMin))-$([string]::Format('{0:N1}', $ProbeTempMax)).",
            ""
        )
        foreach ($item in $newIssues) {
            $probeStr = if ($item.ProbeTemp -ne $null) { [string]::Format("{0:N1}", $item.ProbeTemp) } else { "n/a" }
            $setStr   = if ($item.SetTemp   -ne $null) { [string]::Format("{0:N1}", $item.SetTemp)   } else { "n/a" }
            $bodyLines += @(
                "Output Name: $($item.OutputName)",
                "Device: $($item.DeviceName) (Output #$($item.OutputNumber))",
                "Probe Temp: $probeStr",
                "Set Temp: $setStr",
                "Issue: $($item.Issue)",
                ""
            )
        }
        Send-EmailAlert -Subject $subject -Body ($bodyLines -join [Environment]::NewLine)

        $smsParts = @($newIssues | ForEach-Object {
            "{0} {1} probe {2}" -f $_.DeviceName, $_.OutputName, ($(if ($_.ProbeTemp -ne $null) { [string]::Format("{0:N1}", $_.ProbeTemp) } else { "n/a" }))
        })
        $smsMessage = "Herpstat probe sanity alert: " + ($smsParts -join "; ")
        $null = Send-TextbeltAlert -Message $smsMessage -FailureContext "Probe sanity alert" -AlertCategory 'probe_sanity'
        Write-ActionLog -Message ("Probe sanity issue alert sent for {0} output(s)" -f $newIssues.Count) -Level WARN
    } elseif ($findings.Count -gt 0) {
        Write-ActionLog -Message ("Probe sanity issue still active for {0} output(s); no new alerts sent" -f $findings.Count) -Level WARN
    }

    if ($recoveries.Count -gt 0) {
        $subject = if ($recoveries.Count -eq 1) {
            "RECOVERY: Probe sanity restored for $($recoveries[0].OutputName)"
        } else {
            "RECOVERY: Probe sanity restored ($($recoveries.Count) outputs)"
        }

        $bodyLines = @(
            "One or more probe sanity issues cleared on the latest live reading.",
            ""
        )
        foreach ($item in $recoveries) {
            $probeStr = [string]::Format("{0:N1}", [double]$item.ProbeTemp)
            $setStr   = [string]::Format("{0:N1}", [double]$item.SetTemp)
            $bodyLines += @(
                "Output Name: $($item.OutputName)",
                "Device: $($item.DeviceName) (Output #$($item.OutputNumber))",
                "Probe Temp: $probeStr",
                "Set Temp: $setStr",
                ""
            )
        }
        Send-EmailAlert -Subject $subject -Body ($bodyLines -join [Environment]::NewLine)
        Write-ActionLog -Message ("Probe sanity recovery email sent for {0} output(s)" -f $recoveries.Count)
    }

    $nextState = @{}
    foreach ($key in @($previousState.Keys)) {
        if (-not $currentSeenKeys.ContainsKey($key)) { $nextState[$key] = $true }
    }
    foreach ($key in $currentIssueKeys.Keys) { $nextState[$key] = $true }
    Set-BooleanAlertState -Path $script:probeAlertStatePath -State $nextState
}
function Convert-RowsToHtml {
    param([object[]]$Rows, [datetime]$Since, [datetime]$Until)

    $range = "{0} - {1}" -f $Since.ToString('MM/dd/yyyy hh:mm tt'), $Until.ToString('MM/dd/yyyy hh:mm tt')

    # Filter out nulls and excluded names (defense-in-depth)
    if ($Rows) {
        $Rows = $Rows | Where-Object { $_ -and -not (Test-ExcludeName -Name $_.OutputName) }
    }

    $summaryStats = Get-SummaryOutputStats -Rows $Rows

    if (-not $summaryStats -or $summaryStats.Count -eq 0) {
@"
<html>
  <body style="font-family:Segoe UI,Arial,Helvetica;font-size:14px;">
    <h2>Herpstat Summary</h2>
    <p>Range: $range</p>
    <p><em>No data found in this period.</em></p>
  </body>
</html>
"@
        return
    }

    # ---------- Averages table (one row per device/output) ----------
    $avgRowsHtml = ($summaryStats | ForEach-Object {
        $avgProbeStr = if ($_.AverageProbeTemp -ne $null) { [string]::Format("{0:N1}", $_.AverageProbeTemp) } else { "n/a" }
        $avgSetStr   = if ($_.AverageSetTemp   -ne $null) { [string]::Format("{0:N1}", $_.AverageSetTemp)   } else { "n/a" }
        $deviationStr = if ($_.Deviation -ne $null) { [string]::Format("{0:N1}", $_.Deviation) } else { "n/a" }

@"
<tr>
  <td style='padding:6px;border:1px solid #ddd;'>$($_.DeviceName)</td>
  <td style='padding:6px;border:1px solid #ddd;text-align:center;'>$($_.OutputNumber)</td>
  <td style='padding:6px;border:1px solid #ddd;'>$($_.OutputName)</td>
  <td style='padding:6px;border:1px solid #ddd;text-align:right;'>$avgProbeStr</td>
  <td style='padding:6px;border:1px solid #ddd;text-align:right;'>$avgSetStr</td>
  <td style='padding:6px;border:1px solid #ddd;text-align:right;'>$deviationStr</td>
  <td style='padding:6px;border:1px solid #ddd;text-align:right;'>$($_.ReadingCount)</td>
</tr>
"@
    }) -join "`n"

    $avgTableHtml = @"
<h2>Herpstat Summary</h2>
<p>Range: $range</p>
<h3>Average Temperatures by Output</h3>
<table style="border-collapse:collapse; margin-bottom:18px;">
  <thead>
    <tr>
      <th style='padding:6px;border:1px solid #ddd;'>Device</th>
      <th style='padding:6px;border:1px solid #ddd;'>Output #</th>
      <th style='padding:6px;border:1px solid #ddd;'>Output</th>
      <th style='padding:6px;border:1px solid #ddd;'>Avg Probe Temp</th>
      <th style='padding:6px;border:1px solid #ddd;'>Avg Set Temp</th>
      <th style='padding:6px;border:1px solid #ddd;'>Deviation</th>
      <th style='padding:6px;border:1px solid #ddd;'>Readings</th>
    </tr>
  </thead>
  <tbody>
    $avgRowsHtml
  </tbody>
</table>
"@

    # ---------- Detailed combined table (all rows) ----------
    # Ensure sort keys exist; prefer TimestampISO (_dt populated upstream)
    $sorted = $Rows | Sort-Object -Property _dt, DeviceIP, OutputNumber

    $detailRowsHtml = ($sorted | ForEach-Object {
        $ts  = $_.Timestamp
        $ip  = $_.DeviceIP
        $dev = Get-DeviceFriendlyName -Ip $ip
        $on  = $_.OutputNumber
        $nm  = $_.OutputName
        $pt  = [double]$_.ProbeTemp
        $st  = [double]$_.SetTemp
        $pw  = [int]$_.PowerOutput
@"
<tr>
  <td style='padding:6px;border:1px solid #ddd;white-space:nowrap;'>$ts</td>
  <td style='padding:6px;border:1px solid #ddd;'>$dev</td>
  <td style='padding:6px;border:1px solid #ddd;text-align:center;'>$on</td>
  <td style='padding:6px;border:1px solid #ddd;'>$nm</td>
  <td style='padding:6px;border:1px solid #ddd;text-align:right;'>$([string]::Format("{0:N1}", $pt))</td>
  <td style='padding:6px;border:1px solid #ddd;text-align:right;'>$([string]::Format("{0:N1}", $st))</td>
  <td style='padding:6px;border:1px solid #ddd;text-align:right;'>$pw%</td>
</tr>
"@
    }) -join "`n"

    $detailTableHtml = @"
<h3>All Readings</h3>
<table style="border-collapse:collapse;">
  <thead>
    <tr>
      <th style='padding:6px;border:1px solid #ddd;'>Timestamp</th>
      <th style='padding:6px;border:1px solid #ddd;'>Device</th>
      <th style='padding:6px;border:1px solid #ddd;'>Output #</th>
      <th style='padding:6px;border:1px solid #ddd;'>Name</th>
      <th style='padding:6px;border:1px solid #ddd;'>Probe Temp</th>
      <th style='padding:6px;border:1px solid #ddd;'>Set Temp</th>
      <th style='padding:6px;border:1px solid #ddd;'>Power</th>
    </tr>
  </thead>
  <tbody>
    $detailRowsHtml
  </tbody>
</table>
"@

@"
<html>
  <body style="font-family:Segoe UI,Arial,Helvetica;font-size:14px;">
    $avgTableHtml
    $detailTableHtml
  </body>
</html>
"@
}
function Send-SummaryDeviationAlerts {
    [CmdletBinding()]
    param([object[]]$Rows, [datetime]$Since, [datetime]$Until)

    $summaryStats = Get-SummaryOutputStats -Rows $Rows
    if (-not $summaryStats -or $summaryStats.Count -eq 0) {
        Write-ActionLog -Message "Summary deviation evaluation skipped (no summary stats)"
        return
    }

    $previousState = Get-BooleanAlertState -Path $script:summaryAlertStatePath
    $currentSeenKeys = @{}
    $currentStatsByKey = @{}
    foreach ($item in $summaryStats) {
        $key = Get-OutputAlertKey -DeviceIp ([string]$item.DeviceIP) -OutputNumber ([int]$item.OutputNumber) -OutputName ([string]$item.OutputName)
        Add-Member -InputObject $item -NotePropertyName Key -NotePropertyValue $key -Force
        $currentSeenKeys[$key] = $true
        $currentStatsByKey[$key] = $item
    }

    $flagged = @($summaryStats | Where-Object { $_.DeviationRaw -ne $null -and $_.DeviationRaw -ge $SummaryDeviationThreshold })
    $currentIssueKeys = @{}
    foreach ($item in $flagged) { $currentIssueKeys[$item.Key] = $true }

    $newIssues = @($flagged | Where-Object { -not $previousState.ContainsKey($_.Key) })
    $recoveries = @()
    foreach ($key in @($previousState.Keys)) {
        if ($currentSeenKeys.ContainsKey($key) -and -not $currentIssueKeys.ContainsKey($key)) {
            $recoveries += $currentStatsByKey[$key]
        }
    }

    $range = "{0} - {1}" -f $Since.ToString('MM/dd/yyyy hh:mm tt'), $Until.ToString('MM/dd/yyyy hh:mm tt')
    $thresholdText = [string]::Format("{0:N1}", $SummaryDeviationThreshold)

    if ($newIssues.Count -gt 0) {
        $subject = if ($newIssues.Count -eq 1) {
            "ALERT: Herpstat summary deviation for $($newIssues[0].OutputName)"
        } else {
            "ALERT: Herpstat summary deviations detected ($($newIssues.Count) outputs)"
        }

        $bodyLines = @(
            "One or more outputs met or exceeded the summary deviation threshold of $thresholdText degree(s).",
            "Range: $range",
            ""
        )

        foreach ($item in $newIssues) {
            $avgSetStr   = if ($item.AverageSetTemp   -ne $null) { [string]::Format("{0:N1}", $item.AverageSetTemp)   } else { "n/a" }
            $avgProbeStr = if ($item.AverageProbeTemp -ne $null) { [string]::Format("{0:N1}", $item.AverageProbeTemp) } else { "n/a" }
            $deviationStr = if ($item.Deviation -ne $null) { [string]::Format("{0:N1}", $item.Deviation) } else { "n/a" }

            $bodyLines += @(
                "Output Name: $($item.OutputName)",
                "Device: $($item.DeviceName) (Output #$($item.OutputNumber))",
                "Set Temp: $avgSetStr",
                "Average Temp: $avgProbeStr",
                "Deviation: $deviationStr",
                "Readings: $($item.ReadingCount)",
                ""
            )
        }

        Send-EmailAlert -Subject $subject -Body ($bodyLines -join [Environment]::NewLine)

        $smsParts = @($newIssues | ForEach-Object {
            "{0} {1} set {2} avg {3} dev {4}" -f $_.DeviceName, $_.OutputName, ([string]::Format("{0:N1}", $_.AverageSetTemp)), ([string]::Format("{0:N1}", $_.AverageProbeTemp)), ([string]::Format("{0:N1}", $_.Deviation))
        })
        $smsMessage = "Herpstat summary deviation: " + ($smsParts -join "; ")
        $null = Send-TextbeltAlert -Message $smsMessage -FailureContext "Summary deviation alert" -AlertCategory 'summary_deviation'
        Write-ActionLog -Message ("Summary deviation issue alert sent for {0} output(s)" -f $newIssues.Count) -Level WARN
    } elseif ($flagged.Count -gt 0) {
        Write-ActionLog -Message ("Summary deviation still active for {0} output(s); no new alerts sent" -f $flagged.Count) -Level WARN
    }

    if ($recoveries.Count -gt 0) {
        $subject = if ($recoveries.Count -eq 1) {
            "RECOVERY: Herpstat summary deviation cleared for $($recoveries[0].OutputName)"
        } else {
            "RECOVERY: Herpstat summary deviations cleared ($($recoveries.Count) outputs)"
        }

        $bodyLines = @(
            "One or more outputs returned within the summary deviation threshold of $thresholdText degree(s).",
            "Range: $range",
            ""
        )

        foreach ($item in $recoveries) {
            $avgSetStr   = if ($item.AverageSetTemp   -ne $null) { [string]::Format("{0:N1}", $item.AverageSetTemp)   } else { "n/a" }
            $avgProbeStr = if ($item.AverageProbeTemp -ne $null) { [string]::Format("{0:N1}", $item.AverageProbeTemp) } else { "n/a" }
            $deviationStr = if ($item.Deviation -ne $null) { [string]::Format("{0:N1}", $item.Deviation) } else { "n/a" }

            $bodyLines += @(
                "Output Name: $($item.OutputName)",
                "Device: $($item.DeviceName) (Output #$($item.OutputNumber))",
                "Set Temp: $avgSetStr",
                "Average Temp: $avgProbeStr",
                "Deviation: $deviationStr",
                "Readings: $($item.ReadingCount)",
                ""
            )
        }

        Send-EmailAlert -Subject $subject -Body ($bodyLines -join [Environment]::NewLine)
        Write-ActionLog -Message ("Summary deviation recovery email sent for {0} output(s)" -f $recoveries.Count)
    }

    $nextState = @{}
    foreach ($key in @($previousState.Keys)) {
        if (-not $currentSeenKeys.ContainsKey($key)) { $nextState[$key] = $true }
    }
    foreach ($key in $currentIssueKeys.Keys) { $nextState[$key] = $true }
    Set-BooleanAlertState -Path $script:summaryAlertStatePath -State $nextState

    if ($flagged.Count -eq 0 -and $recoveries.Count -eq 0) {
        Write-ActionLog -Message ("Summary deviation alert not needed (threshold {0:N1})" -f $SummaryDeviationThreshold)
    }
}

# Status email HTML (for ForceStatusNow)
function Convert-RowsToStatusHtml {
    [CmdletBinding()]
    param([object[]]$Rows, [datetime]$At)

    if ($Rows) {
        $Rows = $Rows | Where-Object { $_ }
    }

    if (-not $Rows -or $Rows.Count -eq 0) {
@"
<html>
  <body style="font-family:Segoe UI,Arial,Helvetica;font-size:14px;">
    <h2>Herpstat Status</h2>
    <p>No outputs available at $($At.ToString('MM/dd/yyyy hh:mm tt')).</p>
  </body>
</html>
"@
        return
    }

    $sorted = $Rows | Sort-Object -Property DeviceIP, OutputNumber

    $rowsHtml = ($sorted | ForEach-Object {
        $ip  = $_.DeviceIP
        $dev = Get-DeviceFriendlyName -Ip $ip
        $on  = $_.OutputNumber
        $nm  = $_.OutputName
        $pt  = [double]$_.ProbeTemp
        $st  = [double]$_.SetTemp
        $pw  = [int]$_.PowerOutput
@"
<tr>
  <td style='padding:6px;border:1px solid #ddd;'>$dev</td>
  <td style='padding:6px;border:1px solid #ddd;text-align:center;'>$on</td>
  <td style='padding:6px;border:1px solid #ddd;'>$nm</td>
  <td style='padding:6px;border:1px solid #ddd;text-align:right;'>$([string]::Format("{0:N1}", $pt))</td>
  <td style='padding:6px;border:1px solid #ddd;text-align:right;'>$([string]::Format("{0:N1}", $st))</td>
  <td style='padding:6px;border:1px solid #ddd;text-align:right;'>$pw%</td>
</tr>
"@
    }) -join "`n"

@"
<html>
  <body style="font-family:Segoe UI,Arial,Helvetica;font-size:14px;">
    <h2>Herpstat Status</h2>
    <p>As of $($At.ToString('MM/dd/yyyy hh:mm tt'))</p>
    <table style="border-collapse:collapse;">
      <thead>
        <tr>
          <th style='padding:6px;border:1px solid #ddd;'>Device</th>
          <th style='padding:6px;border:1px solid #ddd;'>Output #</th>
          <th style='padding:6px;border:1px solid #ddd;'>Name</th>
          <th style='padding:6px;border:1px solid #ddd;'>Probe Temp</th>
          <th style='padding:6px;border:1px solid #ddd;'>Set Temp</th>
          <th style='padding:6px;border:1px solid #ddd;'>Power</th>
        </tr>
      </thead>
      <tbody>
        $rowsHtml
      </tbody>
    </table>
  </body>
</html>
"@
}

# Manual summary sender (all devices; does not modify scheduled summary state)
function Send-EmailSummaryNow {
    [CmdletBinding()]
    param([int]$Hours = 12)
    $now   = Get-Date
    $since = $now.AddHours(-[math]::Abs($Hours))
    $rows  = Get-RecentLogRows -Since $since -Until $now
    $html  = [string](Convert-RowsToHtml -Rows $rows -Since $since -Until $now)
    $subj  = "Herpstat Temperature Summary (manual, as of $($now.ToString('MM/dd/yyyy hh:mm tt')))"
    $null  = Send-EmailWithSmsFallback -Subject $subj -Body $html -AsHtml -SmsMessage ("Herpstat manual summary email failed at {0}." -f (Get-SmsSafeTimestamp))
    Write-ActionLog -Message ("Manual summary email sent for range {0} to {1}" -f $since, $now)
}
function Send-EmailSummaryIfWindow {
    $now    = Get-Date
    $state  = Get-LastSummaryInfo
    $target = Get-CurrentTargetTime -Now $now -HourAM $SummaryHourAM -HourPM $SummaryHourPM -MinuteAM $SummaryMinuteAM -MinutePM $SummaryMinutePM
    Write-ActionLog -Message ("Summary target evaluated as {0} with window +/-{1} min" -f $target, $SummaryWindowMinutes)
    if (-not (Test-SummaryWindowEligibility -Now $now -WindowMinutes $SummaryWindowMinutes -Target $target -State $state)) {
        Write-ActionLog -Message "Summary not sent (outside window or already sent)"
        return
    }
    $since = $null
    if ($state -and $state.LastSentISO) { try { $since = [datetime]::Parse([string]$state.LastSentISO) } catch { $since = $null } }
    if (-not $since) { $since = $target.AddHours(-12) }
    Write-ActionLog -Message ("Summary range {0} to {1}" -f $since, $now)
    $rows  = Get-RecentLogRows -Since $since -Until $now
    Write-ActionLog -Message ("Summary rows: {0}" -f ($rows | Measure-Object | Select-Object -ExpandProperty Count))
    $html  = [string](Convert-RowsToHtml -Rows $rows -Since $since -Until $now)
    $subject = "Herpstat Temperature Summary (as of $($now.ToString('MM/dd/yyyy hh:mm tt')))"
    $summarySent = Send-EmailWithSmsFallback -Subject $subject -Body $html -AsHtml -SmsMessage ("Herpstat summary email failed at {0}." -f (Get-SmsSafeTimestamp))
    if ($summarySent) {
        Set-LastSummaryInfo -LastSent $now -LastTarget $target
    } else {
        Write-ActionLog -Message "Summary state not updated because the summary email was not sent successfully" -Level WARN
    }
    Send-SummaryDeviationAlerts -Rows $rows -Since $since -Until $now
}
function Send-TestAlerts {
    [CmdletBinding()]
    param()

    $now12    = Get-Date -Format 'MM/dd/yyyy hh:mm:ss tt'
    $smsStamp = Get-SmsSafeTimestamp

    $issueSubject = "TEST: Herpstat issue alert"
    $issueBody = @(
        "This is a manual Herpstat issue-alert test generated at $now12.",
        "DryRunAlerts: $DryRunAlerts",
        "SMS category: test_issue"
    ) -join [Environment]::NewLine
    Send-EmailAlert -Subject $issueSubject -Body $issueBody
    $null = Send-TextbeltAlert -Message ("Herpstat test issue alert at {0}." -f $smsStamp) -FailureContext "Manual test issue alert" -AlertCategory 'test_issue'

    $recoverySubject = "TEST: Herpstat recovery alert"
    $recoveryBody = @(
        "This is a manual Herpstat recovery-alert test generated at $now12.",
        "Recovery alerts are email-only.",
        "DryRunAlerts: $DryRunAlerts"
    ) -join [Environment]::NewLine
    Send-EmailAlert -Subject $recoverySubject -Body $recoveryBody

    Write-ActionLog -Message "Manual test alerts executed"
}
function Send-TestSummaryDeviationAlerts {
    [CmdletBinding()]
    param()

    $since = (Get-Date).AddHours(-12)
    $until = Get-Date
    $issueRows = @(
        [pscustomobject]@{
            Timestamp    = $since.AddMinutes(10).ToString('MM/dd/yyyy hh:mm:ss tt')
            TimestampISO = $since.AddMinutes(10).ToString('s')
            OutputNumber = 1
            OutputName   = 'Test Summary Habitat'
            ProbeTemp    = 93.8
            SetTemp      = 95.0
            PowerOutput  = 72
            DeviceIP     = '192.168.1.250'
            _dt          = $since.AddMinutes(10)
        },
        [pscustomobject]@{
            Timestamp    = $since.AddHours(2).ToString('MM/dd/yyyy hh:mm:ss tt')
            TimestampISO = $since.AddHours(2).ToString('s')
            OutputNumber = 1
            OutputName   = 'Test Summary Habitat'
            ProbeTemp    = 93.9
            SetTemp      = 95.0
            PowerOutput  = 74
            DeviceIP     = '192.168.1.250'
            _dt          = $since.AddHours(2)
        }
    )
    $recoveryRows = @(
        [pscustomobject]@{
            Timestamp    = $until.AddMinutes(-20).ToString('MM/dd/yyyy hh:mm:ss tt')
            TimestampISO = $until.AddMinutes(-20).ToString('s')
            OutputNumber = 1
            OutputName   = 'Test Summary Habitat'
            ProbeTemp    = 94.7
            SetTemp      = 95.0
            PowerOutput  = 41
            DeviceIP     = '192.168.1.250'
            _dt          = $until.AddMinutes(-20)
        },
        [pscustomobject]@{
            Timestamp    = $until.AddMinutes(-5).ToString('MM/dd/yyyy hh:mm:ss tt')
            TimestampISO = $until.AddMinutes(-5).ToString('s')
            OutputNumber = 1
            OutputName   = 'Test Summary Habitat'
            ProbeTemp    = 94.8
            SetTemp      = 95.0
            PowerOutput  = 38
            DeviceIP     = '192.168.1.250'
            _dt          = $until.AddMinutes(-5)
        }
    )

    $savedPath = $script:summaryAlertStatePath
    if ($DryRunAlerts) {
        $script:summaryAlertStatePath = Join-Path $LogDir "summary_deviation_alerts.test.json"
    }
    try {
        if (-not $DryRunAlerts) {
            Set-BooleanAlertState -Path $script:summaryAlertStatePath -State @{}
        } elseif (Test-Path $script:summaryAlertStatePath) {
            Remove-Item -Path $script:summaryAlertStatePath -Force -ErrorAction SilentlyContinue
        }
        Send-SummaryDeviationAlerts -Rows $issueRows -Since $since -Until $until

        if ($DryRunAlerts) {
            Write-ActionLog -Message "DRY RUN summary deviation recovery leg uses manual email-only test output"
            $subject = "TEST: Herpstat summary deviation recovery"
            $body = @(
                "This is a manual summary deviation recovery test generated at $(Get-Date -Format 'MM/dd/yyyy hh:mm:ss tt').",
                "Output Name: Test Summary Habitat",
                "Set Temp: 95.0",
                "Average Temp: 94.8",
                "Deviation: 0.2",
                "DryRunAlerts: $DryRunAlerts"
            ) -join [Environment]::NewLine
            Send-EmailAlert -Subject $subject -Body $body
        } else {
            Send-SummaryDeviationAlerts -Rows $recoveryRows -Since $since -Until $until
        }
    } finally {
        $script:summaryAlertStatePath = $savedPath
        if ($DryRunAlerts -and (Test-Path (Join-Path $LogDir "summary_deviation_alerts.test.json"))) {
            Remove-Item -Path (Join-Path $LogDir "summary_deviation_alerts.test.json") -Force -ErrorAction SilentlyContinue
        }
    }

    Write-ActionLog -Message "Manual summary deviation test alerts executed"
}
function Send-TestProbeSanityAlerts {
    [CmdletBinding()]
    param()

    $issueRows = @(
        [pscustomobject]@{
            OutputNumber = 1
            OutputName   = 'Test Probe Habitat'
            ProbeTemp    = ($ProbeTempMin - 5)
            SetTemp      = 90.0
            PowerOutput  = 100
            DeviceIP     = '192.168.1.251'
        }
    )
    $recoveryRows = @(
        [pscustomobject]@{
            OutputNumber = 1
            OutputName   = 'Test Probe Habitat'
            ProbeTemp    = (($ProbeTempMin + $ProbeTempMax) / 2)
            SetTemp      = 90.0
            PowerOutput  = 35
            DeviceIP     = '192.168.1.251'
        }
    )

    $savedPath = $script:probeAlertStatePath
    if ($DryRunAlerts) {
        $script:probeAlertStatePath = Join-Path $LogDir "probe_sanity_alerts.test.json"
    }
    try {
        if (-not $DryRunAlerts) {
            Set-BooleanAlertState -Path $script:probeAlertStatePath -State @{}
        } elseif (Test-Path $script:probeAlertStatePath) {
            Remove-Item -Path $script:probeAlertStatePath -Force -ErrorAction SilentlyContinue
        }
        Send-ProbeSanityAlerts -Rows $issueRows

        if ($DryRunAlerts) {
            Write-ActionLog -Message "DRY RUN probe sanity recovery leg uses manual email-only test output"
            $subject = "TEST: Herpstat probe sanity recovery"
            $body = @(
                "This is a manual probe sanity recovery test generated at $(Get-Date -Format 'MM/dd/yyyy hh:mm:ss tt').",
                "Output Name: Test Probe Habitat",
                "Probe Temp: $([string]::Format('{0:N1}', (($ProbeTempMin + $ProbeTempMax) / 2)))",
                "Set Temp: 90.0",
                "DryRunAlerts: $DryRunAlerts"
            ) -join [Environment]::NewLine
            Send-EmailAlert -Subject $subject -Body $body
        } else {
            Send-ProbeSanityAlerts -Rows $recoveryRows
        }
    } finally {
        $script:probeAlertStatePath = $savedPath
        if ($DryRunAlerts -and (Test-Path (Join-Path $LogDir "probe_sanity_alerts.test.json"))) {
            Remove-Item -Path (Join-Path $LogDir "probe_sanity_alerts.test.json") -Force -ErrorAction SilentlyContinue
        }
    }

    Write-ActionLog -Message "Manual probe sanity test alerts executed"
}

# ================= CSV retention and archive =================
function Invoke-LogMaintenance {
    if (-not (Test-Path $script:logPath)) { Write-ActionLog -Message "CSV not found; skipping CSV maintenance"; return }
    $now = Get-Date; $since = $now.AddDays(-$RetentionDays)
    $all = Import-Csv -Path $script:logPath; $keep=@(); $older=@()
    foreach ($r in $all) {
        $dt=$null
        if ($r.PSObject.Properties.Name -contains 'TimestampISO' -and $r.TimestampISO) { try { $dt=[datetime]::Parse([string]$r.TimestampISO) } catch {} }
        if (-not $dt -and $r.Timestamp) { try { $dt=[datetime]::ParseExact([string]$r.Timestamp,'MM/dd/yyyy hh:mm:ss tt',$null) } catch {} }
        if ($dt -and $dt -lt $since) { $older+= $r } else { $keep += $r }
    }
    Write-ActionLog -Message ("CSV maintenance: keep={0} archive={1}" -f $keep.Count,$older.Count)
    if ($older.Count -gt 0) {
        $stamp = (Get-Date -Format 'yyyyMMdd_HHmmss')
        $archiveCsv = Join-Path $ArchiveDir "herpstat_$stamp.csv"
        $older | Select-Object Timestamp,TimestampISO,OutputNumber,OutputName,ProbeTemp,SetTemp,PowerOutput,DeviceIP |
            Export-Csv -Path $archiveCsv -NoTypeInformation -Encoding UTF8
        try {
            Compress-Archive -Path $archiveCsv -DestinationPath ($archiveCsv + '.zip') -Force
            Remove-Item $archiveCsv -Force
            Write-ActionLog -Message "CSV archive zipped and cleaned"
        } catch {
            Write-ActionLog -Message ("CSV archive compression failed: {0}" -f $_.Exception.Message) -Level WARN
        }
    }
    if ($keep.Count -gt 0) {
        $keep | Select-Object Timestamp,TimestampISO,OutputNumber,OutputName,ProbeTemp,SetTemp,PowerOutput,DeviceIP |
            Export-Csv -Path $script:logPath -NoTypeInformation -Encoding UTF8
        Write-ActionLog -Message "CSV rewritten with kept rows"
    } else {
        Remove-Item $script:logPath -Force
        Write-ActionLog -Message "CSV removed (no kept rows)"
    }
}

# ================= Main flow =================
Initialize-VerboseLog
Invoke-VerboseLogRetention

if ($ResetAlertStates) {
    Write-ActionLog -Message "ResetAlertStates requested; clearing saved alert state files"
    Reset-AlertStates
}

$ranManualTests = $false
if ($SendTestAlertsNow) {
    Write-ActionLog -Message "SendTestAlertsNow requested; sending test alerts"
    Send-TestAlerts
    $ranManualTests = $true
}
if ($SendTestSummaryDeviationNow) {
    Write-ActionLog -Message "SendTestSummaryDeviationNow requested; sending summary deviation test alerts"
    Send-TestSummaryDeviationAlerts
    $ranManualTests = $true
}
if ($SendTestProbeSanityNow) {
    Write-ActionLog -Message "SendTestProbeSanityNow requested; sending probe sanity test alerts"
    Send-TestProbeSanityAlerts
    $ranManualTests = $true
}
if ($ranManualTests) {
    Write-ActionLog -Message "Manual test run completed"
    return
}

# Healthchecks start ping
if ($HealthchecksUrl) {
    Invoke-HealthcheckPing -BaseUrl $HealthchecksUrl -Kind start -BodyText ("Start {0}" -f (Get-Date).ToString('MM/dd/yyyy hh:mm:ss tt'))
}

# Optional manual summary first
if ($ForceSummaryNow) {
    Write-ActionLog -Message "ForceSummaryNow requested; sending manual summary first"
    Send-EmailSummaryNow -Hours 12
}

$anyThresholdFailure = $false
$currentRunRows = @()

foreach ($ip in $Devices) {
    $statusUrl = "http://$ip/RAWSTATUS"
    Write-ActionLog -Message ("Begin ping check for {0}" -f $ip)
    $reachable = $false
    try {
        $reachable = Test-Connection -ComputerName $ip -Count $PingCount -Quiet -ErrorAction Stop
        Write-ActionLog -Message ("Ping result for {0}: {1}" -f $ip, ($(if($reachable){"reachable"}else{"unreachable"})))
    } catch {
        Write-ActionLog -Message ("Ping exception for {0}: {1}" -f $ip, $_.Exception.Message) -Level WARN
        $reachable = $false
    }

    if (-not $reachable) {
        $now12    = Get-Date -Format 'MM/dd/yyyy hh:mm:ss tt'
        $smsStamp = Get-SmsSafeTimestamp
        $emailSubject = "ALERT: Herpstat not reachable ($ip)"
        $emailBody    = "Herpstat ($ip) is not responding to ping at $now12."
        $smsMsg       = "Herpstat not responding to ping at $smsStamp."
        Write-ActionLog -Message $emailBody -Level WARN
        $failureResult = Register-DeviceResult -DeviceIp $ip -Success:$false -SmsMessage $smsMsg -FailureContext "Ping failure alert" -EmailSubject $emailSubject -EmailBody $emailBody
        if ($failureResult.ThresholdReached) { $anyThresholdFailure = $true }
        continue
    }

    # Fetch RAWSTATUS
    Write-ActionLog -Message ("Fetching RAWSTATUS from {0}" -f $ip)
    try {
        $resp = Invoke-RestMethod -Uri $statusUrl -Method Get -TimeoutSec 5
        Write-ActionLog -Message ("RAWSTATUS fetch succeeded for {0}" -f $ip)
    } catch {
        $now12    = Get-Date -Format 'MM/dd/yyyy hh:mm:ss tt'
        $smsStamp = Get-SmsSafeTimestamp
        $emailSubject = "ALERT: Failed to read Herpstat status ($ip)"
        $emailBody    = "Failed to read Herpstat status for $ip at $now12. Error: $($_.Exception.Message)"
        $smsMsg       = "Herpstat status read failed at $smsStamp."
        Write-ActionLog -Message $emailBody -Level ERROR
        $failureResult = Register-DeviceResult -DeviceIp $ip -Success:$false -SmsMessage $smsMsg -FailureContext "HTTP status fetch failure" -EmailSubject $emailSubject -EmailBody $emailBody
        if ($failureResult.ThresholdReached) { $anyThresholdFailure = $true }
        continue
    }

    if (-not $resp) {
        $now12    = Get-Date -Format 'MM/dd/yyyy hh:mm:ss tt'
        $smsStamp = Get-SmsSafeTimestamp
        $emailSubject = "ALERT: No data from Herpstat ($ip)"
        $emailBody    = "No data returned by Herpstat status endpoint for $ip at $now12."
        $smsMsg       = "Herpstat status returned no data at $smsStamp."
        Write-ActionLog -Message $emailBody -Level ERROR
        $failureResult = Register-DeviceResult -DeviceIp $ip -Success:$false -SmsMessage $smsMsg -FailureContext "No data from endpoint" -EmailSubject $emailSubject -EmailBody $emailBody
        if ($failureResult.ThresholdReached) { $anyThresholdFailure = $true }
        continue
    }

    # Extract outputs
    Write-ActionLog -Message ("Extracting outputs for {0}" -f $ip)
    $rows = @()
    $rows += Get-HerpstatOutput -obj $resp.output1 -n 1
    $rows += Get-HerpstatOutput -obj $resp.output2 -n 2
    $rows = $rows | Where-Object { $_ -ne $null -and -not (Test-ExcludeName -Name $_.OutputName) }

    if (-not $rows -or $rows.Count -eq 0) {
        $now12    = Get-Date -Format 'MM/dd/yyyy hh:mm:ss tt'
        $smsStamp = Get-SmsSafeTimestamp
        $emailSubject = "ALERT: No outputs in Herpstat response ($ip)"
        $emailBody    = "No outputs found in Herpstat response for $ip at $now12."
        $smsMsg       = "Herpstat returned no outputs at $smsStamp."
        Write-ActionLog -Message $emailBody -Level ERROR
        $failureResult = Register-DeviceResult -DeviceIp $ip -Success:$false -SmsMessage $smsMsg -FailureContext "No outputs in response" -EmailSubject $emailSubject -EmailBody $emailBody
        if ($failureResult.ThresholdReached) { $anyThresholdFailure = $true }
        continue
    }

    # Success path: log and reset failure counter
    Write-ActionLog -Message ("Outputs extracted for {0}: {1}" -f $ip, $rows.Count)
    $devName = Get-DeviceFriendlyName -Ip $ip
    Write-Host "$devName  ($(Get-Date -Format 'MM/dd/yyyy hh:mm:ss tt'))"

    $rows | Format-Table OutputNumber,OutputName,ProbeTemp,SetTemp,PowerOutput -AutoSize
    foreach ($r in $rows) {
        Write-LogRow -OutputNumber $r.OutputNumber -Nickname $r.OutputName -ProbeTemp $r.ProbeTemp -SetTemp $r.SetTemp -PowerOutput $r.PowerOutput -DeviceIp $ip
        # also build in-memory list for optional status email
        $rowMem = [pscustomobject]@{ OutputNumber=$r.OutputNumber; OutputName=$r.OutputName; ProbeTemp=$r.ProbeTemp; SetTemp=$r.SetTemp; PowerOutput=$r.PowerOutput; DeviceIP=$ip }
        $currentRunRows += $rowMem
    }
    Write-ActionLog -Message ("CSV append complete for {0} row(s) on {1}" -f $rows.Count, $ip)
    $null = Register-DeviceResult -DeviceIp $ip -Success:$true -SmsMessage "" -FailureContext ""
}

Write-Host "Logged readings to $script:logPath"
Send-ProbeSanityAlerts -Rows $currentRunRows

# Optional: manual status email of current readings (combined)
if ($ForceStatusNow) {
    $now  = Get-Date
    $html = [string](Convert-RowsToStatusHtml -Rows $currentRunRows -At $now)
    $subj = "Herpstat Status (manual, as of $($now.ToString('MM/dd/yyyy hh:mm tt')))"
    $null = Send-EmailWithSmsFallback -Subject $subj -Body $html -AsHtml -SmsMessage ("Herpstat status email failed at {0}." -f (Get-SmsSafeTimestamp))
    Write-ActionLog -Message "Manual status email sent"
}

Invoke-LogMaintenance
Send-EmailSummaryIfWindow

# Healthchecks success/fail ping based on thresholded device failures
if ($HealthchecksUrl) {
    if ($anyThresholdFailure) {
        Invoke-HealthcheckPing -BaseUrl $HealthchecksUrl -Kind fail -BodyText "One or more devices reached the failure threshold this run"
    } else {
        Invoke-HealthcheckPing -BaseUrl $HealthchecksUrl -Kind success -BodyText ("OK {0}" -f (Get-Date).ToString('MM/dd/yyyy hh:mm:ss tt'))
    }
}

Write-ActionLog -Message ("Run completed. anyThresholdFailure={0}" -f $anyThresholdFailure)
