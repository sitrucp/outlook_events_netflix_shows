# Netflix Data Outlook Event Creation

This repository contains scripts to process Netflix viewing activity data and create Outlook calendar events using the Microsoft Graph API.

## Features
1. **Event Creation**: Create Outlook calendar events directly from Netflix viewing activity data.
2. **Logging**: Maintain logs for processed events and last event creation.

## Workflow Summary
1. Request and download your viewing activity data from Netflix.
2. Use `create_events.py` to create Outlook events from the viewing activity data.
3. Check logs for details and track processing.

## Prerequisites
- **Microsoft Graph API Credentials**: `client_id`, `tenant_id`, `client_secret`, `user_id`
- **Required Libraries**: `pandas`, `numpy`, `requests`, `pytz`

## Usage

### Step 1: Obtain Your Netflix Viewing Activity Data
1. Go to [Netflix Account Settings](https://www.netflix.com/account/getmyinfo).
2. Click **"Download your personal information"**.
3. Wait for an email from Netflix confirming the availability of the data.
4. Download the ZIP file from the provided link.
5. Extract the ZIP file and locate the file at:
   ```
   netflix-report\CONTENT_INTERACTION\ViewingActivity.csv
   ```

### Step 2: Create Outlook Calendar Events
1. Run the `create_events.py` script to process `ViewingActivity.csv` and create Outlook calendar events using the Microsoft Graph API.

## Logs
Two logs are created for tracking progress:
1. **Create Events Log**: Tracks processing details and errors during event creation (`create_events_log.txt`).
2. **Last Event Log**: Tracks the most recent processed event to start from in future runs (`last_event_log.csv`).

## Example Directory Structure
```
repo/
|
├── create_events.py
├── ViewingActivity.csv
│── create_events_log.txt
│── last_event_log.csv
```

## License
This project is licensed under the MIT License. See the LICENSE file for details.
