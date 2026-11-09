# OpenPanel telemetry for the Save As DAISY addin

This folder contains a self-hosted development instance of
[OpenPanel](https://openpanel.dev) to collect telemetry from the Word addin:
usage of the UI actions, conversion counts, addin logs and reported exceptions.

## 1. Start the OpenPanel development instance

Requirements: Docker with the compose plugin.

```bash
cd telemetry/openpanel
cp .env.example .env
# adjust the secrets in .env if you want
docker compose up -d
```

The dashboard is then available at http://localhost:8080 (the reverse proxy is
bound to `127.0.0.1` only, so the instance is not reachable from other
machines).

- First launch: create your admin account from the dashboard
  (`ALLOW_REGISTRATION=true` in the dev `.env`).
- Create a project, then in **Settings > API Clients** create a *write* client
  and copy its `clientId` / `clientSecret`.

## 2. Configure the addin

The telemetry is disabled by default, and enabled by a `telemetry.xml` file in
the addin APPDATA directory (`%APPDATA%\SaveAsDAISY\telemetry.xml`):

```xml
<?xml version="1.0" encoding="utf-8"?>
<telemetry>
  <enabled>true</enabled>
  <!-- URL of the OpenPanel tracking API -->
  <apiUrl>http://localhost:8080/api</apiUrl>
  <clientId>YOUR_CLIENT_ID</clientId>
  <clientSecret>YOUR_CLIENT_SECRET</clientSecret>
</telemetry>
```

A `profileId` (anonymous, stable per Windows user) is generated and stored in
that file on first run, so events are grouped per user in OpenPanel.

## 3. Reported data

| Event             | Reported when                                          | Properties                                        |
| ----------------- | ------------------------------------------------------ | ------------------------------------------------- |
| `addin_action`    | A ribbon / dialog action is used in the addin UI       | `action` (e.g. `SaveAsEPUB3`, `ImportODT`...)     |
| `document_converted` | A document conversion succeeds                      | `script`, `scriptNiceName`                        |
| `addin_log`       | A log entry of level WARN or higher is emitted          | `level`, `message`, exception details            |
| `addin_exception` | An unexpected exception is shown to the user            | `context`, `exceptionType`, `exceptionMessage`, `stackTrace` |

Implementation notes:

- `Common/DaisyConverterLib/Telemetry/OpenPanelTelemetry.cs`: telemetry
  client posting `track` events to the OpenPanel `/track` endpoint
  (`openpanel-client-id` / `openpanel-client-secret` headers).
- `Common/DaisyConverterLib/Telemetry/OpenPanelAppender.cs`: log4net appender
  forwarding the addin logs (threshold: WARN) to OpenPanel.
- `WordAddIn/Connect.cs`: `OpenPanelTelemetry.TrackAction(...)` calls at the
  entry of each ribbon action handler.
- `Common/DaisyConverterLib/Converter.cs`: `TrackConversion(...)` after each
  successful document conversion.
- All telemetry calls are fire-and-forget and silently ignored on failure: a
  missing/unreachable OpenPanel instance never impacts the addin.

## 4. Notes for production

- This compose file is a development instance (loopback only, registration
  open, dev secrets). For a real deployment follow the official OpenPanel
  self-hosting guide: https://openpanel.dev/docs/self-hosting/self-hosting
- Consider informing users that usage data is collected (the telemetry can be
  disabled by removing the `telemetry.xml` file or setting `<enabled>false</enabled>`).
