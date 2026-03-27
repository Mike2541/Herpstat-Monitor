# Herpstat-Monitor

PowerShell monitor for Herpstat thermostats with CSV logging, Gmail API email alerts, Textbelt SMS alerts, summary reporting, probe sanity checks, recovery notifications, and Healthchecks.io integration.

## Quick Start

1. Open `HerpstatMonitor.ps1`.
2. Fill in the user configuration values near the top:
   - `Devices` and `DeviceNames`
   - `MailFrom` and `MailTo`
   - `GoogleOAuthClientId`
   - `GoogleOAuthClientSecret`
   - `GoogleOAuthRefreshToken`
   - optional `SmsTo`, `TextbeltApiKey`, and `HealthchecksUrl`
3. Run a dry-run test:

```powershell
.\HerpstatMonitor.ps1 -SendTestAlertsNow -DryRunAlerts
```

4. Run a real alert test:

```powershell
.\HerpstatMonitor.ps1 -SendTestAlertsNow
```

## Common Test Commands

Basic alert flow:

```powershell
.\HerpstatMonitor.ps1 -SendTestAlertsNow -DryRunAlerts
.\HerpstatMonitor.ps1 -SendTestAlertsNow
```

Summary deviation test:

```powershell
.\HerpstatMonitor.ps1 -SendTestSummaryDeviationNow -DryRunAlerts
```

Probe sanity test:

```powershell
.\HerpstatMonitor.ps1 -SendTestProbeSanityNow -DryRunAlerts
```

Reset saved alert states:

```powershell
.\HerpstatMonitor.ps1 -ResetAlertStates -DryRunAlerts
```

Manual status or summary tests without device access:

```powershell
.\HerpstatMonitor.ps1 -ForceStatusNow -Devices @() -SkipHealthchecks
.\HerpstatMonitor.ps1 -ForceSummaryNow -Devices @() -SkipHealthchecks
```

## Notes

- Runtime logs and state files default to `Desktop\Herpstat`, not the repo folder.
- The script is set up for top-of-file configuration so it is easier for non-technical users to edit.
- Before sharing your own configured copy, rotate any real OAuth, SMS, or Healthchecks secrets.
