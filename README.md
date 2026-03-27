# Herpstat-Monitor

PowerShell monitor for Herpstat thermostats with CSV logging, Gmail API email alerts, Textbelt SMS alerts, summary reporting, probe sanity checks, recovery notifications, and Healthchecks.io integration.

## What You Need

- Windows with PowerShell
- Access to your Herpstat device IPs on your local network
- A Gmail account for email alerts
- A Google Cloud project with Gmail API enabled
- Optional: a Textbelt API key for SMS alerts
- Optional: a Healthchecks.io check for run monitoring

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

## Configuration Guide

This script is intentionally set up so most users can configure it by editing values near the top of `HerpstatMonitor.ps1`. You do not need to use environment variables for the public template.

### Device Settings

Fill in these values first:

- `Devices`
  Add your Herpstat IP addresses.
- `DeviceNames`
  Map each IP to a friendly name like `Herpstat1` or `Rack Left`.
- `MailFrom`
  The Gmail address the script will send from.
- `MailTo`
  The email address that should receive alerts and summaries.

Example:

```powershell
[string[]]$Devices = @("192.168.1.50","192.168.1.51"),
[hashtable]$DeviceNames = @{
    '192.168.1.50' = 'Herpstat1'
    '192.168.1.51' = 'Herpstat2'
},
```

### Gmail API Setup

The script uses the Gmail API with OAuth refresh-token flow. You need these three values:

- `GoogleOAuthClientId`
- `GoogleOAuthClientSecret`
- `GoogleOAuthRefreshToken`

You do not need to store the short-lived access token in the script.

#### Step 1: Create or choose a Google Cloud project

Open the Google Cloud Console and create a project or select an existing one.

#### Step 2: Enable the Gmail API

Enable the Gmail API for that project:

https://console.cloud.google.com/apis/library/gmail.googleapis.com

#### Step 3: Configure the Google Auth platform

Open the Google Auth platform / OAuth consent configuration and fill in the basic app information.

If you are using a personal Gmail account:

1. Set the app audience to `External`.
2. If the app is still in `Testing`, add your Gmail address as a test user.

If you skip the test-user step, you can get:

`Error 403: access_denied`

#### Step 4: Create OAuth client credentials

Create an OAuth client credential as a `Web application`.

Add this exact authorized redirect URI:

`https://developers.google.com/oauthplayground`

After saving, copy:

- Client ID
- Client Secret

#### Step 5: Generate a refresh token

Open the OAuth Playground:

https://developers.google.com/oauthplayground/

Then:

1. Click the gear icon.
2. Enable `Use your own OAuth credentials`.
3. Paste your Client ID and Client Secret.
4. Use this scope:

`https://www.googleapis.com/auth/gmail.send`

5. Click `Authorize APIs`.
6. Sign in with the Gmail account you want the script to send from.
7. Click `Exchange authorization code for tokens`.
8. Copy the `refresh token`.

Paste the following into `HerpstatMonitor.ps1`:

- `GoogleOAuthClientId`
- `GoogleOAuthClientSecret`
- `GoogleOAuthRefreshToken`

#### Gmail API Notes

- `MailFrom` should be the Gmail account you authorized, or a valid Gmail "send as" identity on that account.
- The script refreshes its own access tokens automatically. Only the refresh token needs to be stored.
- For a private test setup, keeping these values in the script is fine. For production or shared environments, move them to environment variables or a secure secret store.

### Textbelt SMS Setup

SMS is optional. If you leave `SmsTo` or `TextbeltApiKey` blank, SMS will stay disabled.

To set up Textbelt:

1. Create or buy an API key at Textbelt.
2. Put your destination phone number in `SmsTo`.
3. Put your API key in `TextbeltApiKey`.

Useful notes:

- In the U.S. and Canada, a normal 10-digit number usually works.
- Outside the U.S., E.164 format is the safest choice.
- Textbelt supports a free test key for limited use, and its docs also describe appending `_test` to your key to validate requests without consuming quota.

### Healthchecks.io Setup

Healthchecks is optional. If you leave `HealthchecksUrl` blank, Healthchecks pings are skipped.

To set it up:

1. Create a check in Healthchecks.io.
2. Copy the ping URL for that check.
3. Paste it into `HealthchecksUrl`.

This script sends:

- `/start` when the run begins
- base URL on success
- `/fail` when a device reaches the failure threshold

If you want to skip Healthchecks temporarily during testing, use:

```powershell
.\HerpstatMonitor.ps1 -SkipHealthchecks
```

### Recommended First-Time Test Order

1. Fill in Gmail API values and email addresses.
2. Run:

```powershell
.\HerpstatMonitor.ps1 -SendTestAlertsNow -DryRunAlerts
```

3. Run:

```powershell
.\HerpstatMonitor.ps1 -SendTestAlertsNow
```

4. If you plan to use SMS, configure Textbelt and repeat the same test.
5. If you plan to use Healthchecks, add the ping URL and test a normal run.

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

## Helpful Links

- Gmail API sending guide: https://developers.google.com/workspace/gmail/api/guides/sending
- Google OAuth consent configuration: https://developers.google.com/workspace/guides/configure-oauth-consent
- Google credential creation guide: https://developers.google.com/workspace/guides/create-credentials
- OAuth Playground setup notes: https://developers.google.com/google-ads/api/docs/oauth/playground
- Textbelt docs: https://docs.textbelt.com/
- Healthchecks.io HTTP pinging API: https://healthchecks.io/docs/http_api/
