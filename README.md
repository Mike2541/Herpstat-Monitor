# Herpstat-Monitor

PowerShell monitor for Herpstat thermostats with CSV logging, Gmail API email alerts, Textbelt SMS alerts, summary reporting, probe sanity checks, recovery notifications, and Healthchecks.io integration.

## What You Need

- Windows with PowerShell
- A supported Spyder Robotics Herpstat SpyderWeb device on your local network
- A Gmail account for email alerts
- A Google Cloud project with Gmail API enabled
- Optional: a Textbelt API key for SMS alerts
- Optional: a Healthchecks.io check for run monitoring

## Supported Devices

This script is intended for Spyder Robotics Herpstat models that include the SpyderWeb Wi-Fi/web interface.

Confirmed/documented models:

- Herpstat 1 SpyderWeb
- Herpstat 2 SpyderWeb
- Herpstat 4 SpyderWeb
- Herpstat 6 SpyderWeb

It relies on the local SpyderWeb web interface and the `RAWSTATUS` endpoint used by these Wi-Fi-enabled models. Older non-SpyderWeb models are not the target for this script.

Manuals:

- Herpstat 1/2 SpyderWeb manual: https://www.spyderrobotics.com/manuals/Herpstat12_SpyderWeb_manual.pdf
- Herpstat 4/6 SpyderWeb manual: https://www.spyderrobotics.com/manuals/Herpstat46_SpyderWeb_manual.pdf

## Why Use This Script Instead of Only the Built-In Features?

Spyder Robotics already includes useful built-in SpyderWeb features, and for many users those may be enough.

Documented built-in features include:

- local web status viewing
- a history graph in the web interface
- scheduled email status updates
- emergency email alerts for conditions such as probe errors, high/low alarms, and device resets
- missed scheduled upload email alerts when the device fails to upload to the SpyderWeb site
- optional upload of status and charts to `herpstat.com` for online viewing
- advanced status integration options including a custom upload target and the `RAWSTATUS` page

This script is aimed at advanced users who want more control over how monitoring and reporting work.

Reasons to use it:

- combine multiple Herpstat devices into one reporting workflow
- keep local CSV history for your own records
- send summary emails based on averages across a time window instead of only a point-in-time snapshot
- choose your own schedule through Windows Task Scheduler
- keep monitoring independent of the SpyderWeb cloud upload path
- use custom issue handling when a built-in missed-upload notice may only send once until the device uploads successfully again
- add issue and recovery workflows that fit your own preferences
- add optional SMS alerts through Textbelt
- add external run monitoring through Healthchecks.io
- customize thresholds, testing behavior, and alert timing more than the stock web interface allows

In short:

- if the built-in SpyderWeb features already cover your needs, they are a good and free option from Spyder Robotics
- if you want more customizable reporting and alerting, this script is the advanced-user layer on top

## Quick Start

1. Open `HerpstatMonitor.ps1`.
2. Fill in the user configuration values near the top:
   - `Devices` and `DeviceNames`
   - `MailFrom` and `MailTo`
   - `GoogleOAuthClientId`
   - `GoogleOAuthClientSecret`
   - `GoogleOAuthRefreshToken`
   - optional `SmsTo`, `TextbeltApiKey`, and `HealthchecksUrl`
3. Validate the configuration without sending alerts:

```powershell
.\HerpstatMonitor.ps1 -SendTestAlertsNow -DryRunAlerts
```

4. Validate live alert delivery:

```powershell
.\HerpstatMonitor.ps1 -SendTestAlertsNow
```

5. Set up Windows Task Scheduler for normal operation.

## Configuration Guide

This script is intentionally set up so most users can configure it by editing values near the top of `HerpstatMonitor.ps1`. You do not need to use environment variables for the public template.

### Device Settings

Fill in these values first:

- `Devices`
  Add your Herpstat IP addresses or resolvable hostnames.
- `DeviceNames`
  Map each device entry to a friendly name like `Herpstat1` or `Rack Left`.
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

### IP Address vs Hostname

The script uses each `Devices` entry directly for:

- `Test-Connection`
- `http://<device>/RAWSTATUS`

That means a resolvable hostname can work, not just a numeric IP address.

In practice, static or reserved IPs are the recommended setup.

Both SpyderWeb manuals state that the Herpstat initially receives a dynamic IP from the router and that it is often better to reserve a fixed/static IP so the address does not change over time. That recommendation is especially helpful for this script because scheduled monitoring is more reliable when the device address stays the same.

If you do use a hostname instead of an IP:

- make sure Windows can resolve it reliably
- use that exact same hostname string as the key in `DeviceNames`

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

### Initial Validation

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
5. If you plan to use Healthchecks, add the ping URL and then switch to scheduled runs.

## Windows Task Scheduler Setup

This script is intended to run automatically on a schedule. The one-off commands in this README are mainly for setup, validation, and troubleshooting. For normal day-to-day operation, run it with Windows Task Scheduler.

### Choosing a Schedule

This script works with either lighter scheduled runs or more frequent polling.

Common patterns:

- `Every hour on the hour`
  Good for lighter monitoring when slower detection is acceptable.
- `Every 5 to 15 minutes`
  Better if you want faster alerts and more tolerance for schedule drift.

Tradeoffs:

- More frequent runs detect device failures, probe sanity issues, and recoveries sooner.
- More frequent runs also make it easier to hit summary windows even if the task does not fire at the exact minute.
- Hourly on-the-hour runs are completely valid, especially if your summary times are also on the hour.

If you run hourly on the hour:

- `SummaryWindowMinutes = 20` is usually fine for `8:00 AM` and `8:00 PM` summaries.
- `FailureThreshold = 2` means a device issue alert will normally happen after about 2 hours of consecutive failures.

If your task timing drifts or you want faster alerting, either widen `SummaryWindowMinutes` or run the task more frequently.

### Create the Task

1. Open `Task Scheduler`.
2. Click `Create Task`.
3. On the `General` tab:
   - give it a name like `Herpstat Monitor`
   - select `Run whether user is logged on or not` if you want it to keep working in the background
   - enable `Run with highest privileges` only if your environment needs it
4. On the `Triggers` tab:
   - create a new trigger
   - choose `Daily`
   - check `Repeat task every`
   - set it to your preferred interval such as `5 minutes`, `15 minutes`, or `1 hour`
   - set `for a duration of` to `Indefinitely`
5. On the `Actions` tab:
   - create a new action
   - `Program/script`:

```text
powershell.exe
```

   - `Add arguments`:

```text
-ExecutionPolicy Bypass -File "C:\Path\To\HerpstatMonitor.ps1"
```

   - `Start in`:

```text
C:\Path\To
```

6. On the `Conditions` tab:
   - disable `Start the task only if the computer is on AC power` if this is a laptop and you want it to run on battery
   - enable `Wake the computer to run this task` if needed
7. On the `Settings` tab:
   - enable `Allow task to be run on demand`
   - enable `Run task as soon as possible after a scheduled start is missed`
   - set `If the task is already running` to `Do not start a new instance`

### Example

If the script is stored in:

```text
C:\Users\YourName\Documents\Herpstat-Monitor\HerpstatMonitor.ps1
```

then use:

`Program/script`

```text
powershell.exe
```

`Add arguments`

```text
-ExecutionPolicy Bypass -File "C:\Users\YourName\Documents\Herpstat-Monitor\HerpstatMonitor.ps1"
```

`Start in`

```text
C:\Users\YourName\Documents\Herpstat-Monitor
```

### Scheduler Verification

After saving the task:

1. Right-click the task and choose `Run`.
2. Check the newest file in `Desktop\Herpstat\Verbose`.
3. Confirm it created or updated:
   - the verbose log
   - the CSV log
   - any test emails or alerts you expected

If you want to validate the scheduled task without sending live alerts first, temporarily use:

```text
-ExecutionPolicy Bypass -File "C:\Path\To\HerpstatMonitor.ps1" -DryRunAlerts -SkipHealthchecks
```

### Scheduler Tips

- Run the task on a machine that can actually reach your Herpstat IP addresses.
- If the computer sleeps often, consider wake settings or a device that stays on all the time.
- If you change the script path later, update the scheduled task action too.
- If Task Scheduler says the task ran but nothing happened, the verbose log in `Desktop\Herpstat\Verbose` is the first place to check.

## Validation and Troubleshooting Commands

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

Optional one-off status or summary checks without device access:

```powershell
.\HerpstatMonitor.ps1 -ForceStatusNow -Devices @() -SkipHealthchecks
.\HerpstatMonitor.ps1 -ForceSummaryNow -Devices @() -SkipHealthchecks
```

## Troubleshooting

### Gmail API

`Error 403: access_denied`

- Your Google OAuth app is probably still in `Testing` and your Gmail account is not listed as a test user.
- Add the Gmail account under the app's test users and try again.

`invalid_grant`

- The refresh token is invalid, expired, revoked, or tied to a different OAuth client.
- Generate a new refresh token in OAuth Playground using the same client ID and client secret stored in the script.

`redirect_uri_mismatch`

- The OAuth client is missing the OAuth Playground redirect URI.
- Make sure the client includes:
  `https://developers.google.com/oauthplayground`

Email sends fail even though OAuth is configured

- Confirm `MailFrom` matches the Gmail account you authorized, or a valid Gmail "send as" alias on that account.
- Run:

```powershell
.\HerpstatMonitor.ps1 -SendTestAlertsNow
```

- Then check the newest log in `Desktop\Herpstat\Verbose`.

### Textbelt SMS

SMS never sends

- Confirm both `SmsTo` and `TextbeltApiKey` are filled in.
- Check the verbose log for the Textbelt API response.
- If you are outside the U.S., try E.164 phone format.

SMS is being skipped

- The script uses per-category SMS cooldowns.
- A recent alert in the same category can suppress another SMS until the cooldown expires.
- Check the verbose log for `SMS suppressed by rate limit`.

SMS testing without consuming quota

- Textbelt documents a free test key and also supports appending `_test` to your key for test requests.
- That is useful when validating formatting before using live SMS credits.

### Device Connectivity

Devices are always unreachable

- Confirm the IPs in `Devices` are correct.
- Confirm the Windows machine running the script is on the same network and can reach the devices.
- Try pinging the Herpstat IP manually from the same machine.
- If the script is running on a different PC than usual, local firewall or routing may be different.

No outputs are found

- The device responded, but the script did not get usable output objects.
- Check the verbose log to see whether the RAWSTATUS request succeeded and whether output names were excluded.
- Remember that names like `Nickname` or `Nickname2` are intentionally ignored by the script.

### Summary and Alert Timing

Summary email did not send

- The script only sends the scheduled summary inside the configured `SummaryWindowMinutes`.
- It also records the last sent target time so it does not resend the same scheduled summary repeatedly.
- If you run the script hourly, running it on the hour lines up best with on-the-hour summary targets.
- If your task timing drifts, widen `SummaryWindowMinutes` or run the task more frequently.
- Check:
  - `SummaryHourAM`
  - `SummaryHourPM`
  - `SummaryWindowMinutes`
  - `last_summary.json` in `Desktop\Herpstat`

Summary deviation or probe sanity alert did not repeat

- These alerts are stateful.
- Once an issue is active, the script avoids sending the same first-occurrence alert over and over.
- To retest first-occurrence behavior, use:

```powershell
.\HerpstatMonitor.ps1 -ResetAlertStates -DryRunAlerts
```

Recovery alert did not send SMS

- Recovery alerts are intentionally email-only.
- SMS is reserved for issue alerts.

### Task Scheduler

Task Scheduler says the task ran, but nothing happened

- Check the newest log in `Desktop\Herpstat\Verbose`.
- Confirm the scheduled task `Start in` folder is correct.
- Confirm the script path in `-File` is correct.
- Make sure the scheduled machine can still reach your Herpstat IPs.

The task works manually but not on schedule

- Make sure `Run whether user is logged on or not` is configured if needed.
- Check the `Conditions` tab for power or sleep restrictions.
- Confirm the task is repeating at your intended interval for `Indefinitely`.

### Fastest Debug Path

If something is not behaving the way you expect:

1. Run a dry-run validation command.
2. Run the matching live validation command if needed.
3. Open the newest file in `Desktop\Herpstat\Verbose`.
4. Check the JSON state files in `Desktop\Herpstat` if alert timing or repeat behavior seems wrong.

## Notes

- Runtime logs and state files default to `Desktop\Herpstat`, not the repo folder.
- The script is set up for top-of-file configuration so it is easier for non-technical users to edit.
- Normal operation is intended to be scheduled through Windows Task Scheduler.
- Before sharing your own configured copy, rotate any real OAuth, SMS, or Healthchecks secrets.

## Helpful Links

- Gmail API sending guide: https://developers.google.com/workspace/gmail/api/guides/sending
- Google OAuth consent configuration: https://developers.google.com/workspace/guides/configure-oauth-consent
- Google credential creation guide: https://developers.google.com/workspace/guides/create-credentials
- OAuth Playground setup notes: https://developers.google.com/google-ads/api/docs/oauth/playground
- Textbelt docs: https://docs.textbelt.com/
- Healthchecks.io HTTP pinging API: https://healthchecks.io/docs/http_api/
