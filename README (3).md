# Route Status Automation Toolkit

Two Tampermonkey userscripts that work together to automate route-status tracking on a delivery-routing web portal and streamline getting that data into an Excel Online workbook — replacing what used to be a repetitive manual copy/paste/lookup process with a one-click pipeline.

> **Note:** This is a sanitized, portfolio/demo build. It was originally built for a specific internal logistics workflow; all company-specific URLs, site identifiers, and naming conventions have been replaced with generic placeholders. To use this against a real system, you'll need to point the `@match` directives and a few config constants (flagged in the code) at your own target site.

## Screenshots

**Routing portal + live automation panel**
The floating panel in the top-right tracks route processing in real time and exports a status grid straight to CSV.

![Routing engine automation panel](./screenshots/routing-engine-panel.png)

**Excel Online companion — exception grid + CSV injector**
Paste in a truck/team assignment list, override individual stop codes by hand when needed, then stream the formatted result straight to the clipboard for pasting into the spreadsheet.

![Excel matrix companion panel](./screenshots/excel-matrix-companion-panel.png)

**Installed as standard Tampermonkey userscripts**
No browser extension packaging required — just two scripts running through Tampermonkey.

![Tampermonkey installed scripts](./screenshots/tampermonkey-installed-scripts.png)

## What's in here

### 1. `route-status-automation.user.js`
Runs on a delivery-routing portal. Given a list of route IDs, it automatically:
- Switches location/context per route
- Opens each route's stop list and reads the status of every stop (delivered, in transit, arrived, not loaded, etc.)
- Normalizes messy on-screen statuses into short, consistent codes
- Flags stops that need a manual check instead of guessing
- Renders a live status grid directly in the browser
- Exports everything to a CSV file, ready for the second script
- Automatically resumes its polling cycle after a scheduled page refresh — no manual restart needed
- Supports multiple saved tabs/workspaces for tracking several route queues at once

### 2. `excel-matrix-companion.user.js`
Runs on Excel Online / SharePoint. Takes the CSV from script 1 and:
- Matches each route to a truck/team assignment (pasted in once, saved locally)
- Lets you manually override specific stop codes through a point-and-click grid (20 stops per route)
- Auto-corrects known ambiguous codes into a consistent format
- Formats everything into a tab-separated payload and copies it straight to your clipboard, ready to paste into the spreadsheet
- Supports multiple saved tabs/workspaces for different configurations

## Why this exists

Manually checking and re-typing route/stop statuses one at a time took roughly 12–13 minutes per check on average. This automation brings that down to roughly 8 minutes — a **~35% time reduction** and a **1.5x speed-up** — while also eliminating the transcription errors that showed up in manual runs (observed 50% error rate manually vs. 0% automated, across a small sample). Scaled across a full-size logistics coordination department, that reclaimed time adds up to the equivalent of several additional full-time coordinators' worth of capacity per year, freed up for higher-value work like customer handling and callbacks instead of manual status monitoring.

## Setup

1. Install [Tampermonkey](https://www.tampermonkey.net/) in your browser.
2. Create a new userscript for each file and paste in the contents.
3. Update the `@match` line at the top of each script to point at your own target site(s).
4. In `route-status-automation.user.js`, update:
   - `ROUTE_ID_PATTERN` — the regex that identifies route IDs in your data
   - `STORE_KEYWORD_REGEX` — the keyword your system uses to flag store-stops vs. customer stops
5. Save, refresh the target page, and the control panel should appear automatically.

## Disclaimer

This code is shared for portfolio and educational purposes. It automates interactions with a web page's DOM (clicking buttons, reading rendered text, filling inputs) and does not use or expose any private API, credentials, or backend access. Before running any browser automation against a work system, confirm it's permitted under your employer's IT and acceptable-use policies.

## License

MIT — see [LICENSE](LICENSE) for details.

## Author

Danilo Jose Rengifo Sulbaran
