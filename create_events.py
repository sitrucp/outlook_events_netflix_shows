import pandas as pd
import requests
from datetime import datetime, timedelta
import time
import pytz
import sys 
import os


#--- Get msgraph config variables ---#
config_msgraph_path = os.getenv("ENV_VARS_PATH")  # Get path to directory contaiining config_msgraph.py
if not config_msgraph_path:
    raise ValueError("ENV_VARS_PATH environment variable not set")
sys.path.insert(0, config_msgraph_path)

from config_msgraph import config_msgraph

client_id=config_msgraph["client_id"]
tenant_id=config_msgraph["tenant_id"]
client_secret=config_msgraph["client_secret"]
user_id=config_msgraph["user_id"]

input_file = "ViewingActivity.csv"
output_file = "FilteredViewingActivity.csv"
log_file = "last_event_date.csv"

country_timezones = {
    'CA': 'America/Toronto',  # Canada, Toronto
    'US': 'America/New_York',  # United States, New York
    'SG': 'Asia/Singapore',  # Singapore
    'MY': 'Asia/Kuala_Lumpur',  # Malaysia, Kuala Lumpur
    'NL': 'Europe/Amsterdam',  # Netherlands, Amsterdam
    'PL': 'Europe/Warsaw',  # Poland, Warsaw
    'DE': 'Europe/Berlin',  # Germany, Berlin
    'HR': 'Europe/Zagreb',  # Croatia, Zagreb
    'GR': 'Europe/Athens',  # Greece, Athens
    'HU': 'Europe/Budapest',  # Hungary, Budapest
    'FR': 'Europe/Paris',  # France, Paris
    'AE': 'Asia/Dubai',  # United Arab Emirates, Dubai
    'SE': 'Europe/Stockholm',  # Sweden, Stockholm
    'JP': 'Asia/Tokyo',  # Japan, Tokyo
    'PT': 'Europe/Lisbon',  # Portugal, Lisbon
    'GB': 'Europe/London',  # United Kingdom, London
    'SA': 'Asia/Riyadh',  # Saudi Arabia, Riyadh
    'CZ': 'Europe/Prague',  # Czech Republic, Prague
    'DK': 'Europe/Copenhagen',  # Denmark, Copenhagen
    'EE': 'Europe/Tallinn',  # Estonia, Tallinn
    'FI': 'Europe/Helsinki',  # Finland, Helsinki
    'NO': 'Europe/Oslo',  # Norway, Oslo
}

# Filter the data converting '00:00:05' format for minutes filter
def filter_duration(duration_str):
    try:
        hours, minutes, seconds = duration_str.split(':') 
        return int(hours) * 60 + int(minutes)  # Returns total minutes
    except ValueError as e:
        print(f"Error parsing duration: {e}")
        return 0

def convert_to_local_time(utc_time, country_code):
    """Convert UTC datetime to local time based on country code, accounting for DST, and return the timezone."""
    timezone_str = country_timezones.get(country_code)
    if timezone_str:
        # Ensure the datetime is timezone-aware
        utc_zone = pytz.utc
        utc_time = utc_time.replace(tzinfo=utc_zone)
        
        # Convert to the target timezone with DST consideration
        target_timezone = pytz.timezone(timezone_str)
        local_time = utc_time.astimezone(target_timezone)
        return local_time, timezone_str  # Return both the local time and the timezone string
    else:
        return utc_time, "UTC"  # Fallback to UTC if no timezone is found

def get_country_code(country_str):
    """Extracts the country code from the 'Country' column."""
    return country_str.split(' ')[0]  # Assumes format "Code (Country Name)"

# New function to apply the conversion and capture both local time and timezone
def apply_conversion_and_capture_timezone(row):
    local_time, timezone_str = convert_to_local_time(row['Start Time'], get_country_code(row['Country']))
    return pd.Series([local_time, timezone_str], index=['Local Start Time', 'Timezone'])

#--- Function to obtain an access token ---#

def get_access_token(client_id, tenant_id, client_secret):
    print("Starting token retrieval process...")
    url = f'https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token'
    print(f"Token URL: {url}")

    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    print("Request headers:", headers)

    data = {
        'client_id': client_id,
        'scope': 'https://graph.microsoft.com/.default',
        'client_secret': client_secret,
        'grant_type': 'client_credentials',
    }
    print("Request data:", {k: v if k != 'client_secret' else '[REDACTED]' for k, v in data.items()})
    print("Sending POST request to token endpoint...")

    response = requests.post(url, headers=headers, data=data)
    print(f"Response status code: {response.status_code}")

    if response.status_code == 200:
        print("Token retrieved successfully")
        token = response.json().get('access_token')
        print(f"Access token: {token[:10]}...{token[-10:]}") # Print first and last 10 characters of token
    else:
        print("Failed to retrieve token")
        print("Response content:", response.text)
        
    response.raise_for_status()  # Raises an exception for HTTP error codes
    return response.json().get('access_token')

#--- Function to create a calendar event using Microsoft Graph API ---#

def create_calendar_event(access_token, row):
    # Calculate the local end time by adding the duration to the local start time
    duration_hours, duration_minutes, _ = [int(x) for x in row['Duration'].split(':')]
    duration_delta = timedelta(hours=duration_hours, minutes=duration_minutes)
    local_end_time = row['Local Start Time'] + duration_delta

    # Format the start and end times for the event payload
    start_time_formatted = row['Local Start Time'].strftime('%Y-%m-%dT%H:%M:%S')
    end_time_formatted = local_end_time.strftime('%Y-%m-%dT%H:%M:%S')

    # Format the start time for the description using the original local time
    start_time_for_description = row['Local Start Time'].strftime('%Y-%m-%d %H:%M:%S')
    end_time_for_description = local_end_time.strftime('%Y-%m-%d %H:%M:%S')

    # create the event description
    description_html = (
        f"Title: {row['Title']}<br>"
        f"Start: {start_time_for_description}<br>"
        f"End: {end_time_for_description}<br>"
        f"Duration: {row['Duration']}<br>"
        f"Attributes: {(str(row['Attributes']).replace(',', ', ') if pd.notna(row['Attributes']) else 'None')}<br>"
        f"Device: {row['Device Type']}<br>"
        f"Country: {row['Country']}"
    )

    # Then, include this HTML-formatted description in your payload
    event_payload = {
        "subject": f"Netflix: {row['Title']}",
        "start": {
            "dateTime": start_time_formatted,
            "timeZone": row['Timezone']
        },
        "end": {
            "dateTime": end_time_formatted,
            "timeZone": row['Timezone']
        },
        "body": {
            "contentType": "HTML",
            "content": description_html
        },
        "categories": ["Netflix"]
    }

    # Send the request to create the event
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
    print("Sending request to create event...")  # Debug print statement
    response = requests.post(f"https://graph.microsoft.com/v1.0/users/{user_id}/events",
                             headers=headers, json=event_payload)
    print(f"Response status code for event creation: {response.status_code}")  # Debug print statement
    response.raise_for_status()  # Ensure successful request

#--- Main script to process the CSV data and create events ---#

def main():
    print("Starting main function...")  # Debug print statement
    
    # Obtain access token
    access_token = get_access_token(client_id, tenant_id, client_secret)
    print('Access token obtained successfully')
    
    # Read the last record date from the log file
    print(f"Reading log file: {log_file}")  # Debug print statement
    log_df = pd.read_csv(log_file)
    last_record_date_str = log_df.iloc[0]['last_record_date']  # Assuming there's only one record
    print(f"Last record date string: {last_record_date_str}")  # Debug print statement
    last_record_date = pd.to_datetime(last_record_date_str).date()
    print(f"Last record date: {last_record_date}")  # Debug print statement

    # Read the CSV file, ensuring datetime parsing
    print(f"Reading input file: {input_file}")  # Debug print statement
    df = pd.read_csv(input_file, parse_dates=["Start Time"])
    print("CSV file read successfully")  # Debug print statement

    # Filter to exclude non-relevant records
    df_filtered = df[(df["Supplemental Video Type"].isnull()) & 
                              (df['Duration'].apply(filter_duration) >= 10)].copy()
    print(f"Filtered DataFrame:\n{df_filtered.head()}")  # Debug print statement
    
    # Convert 'Start Time' to local time and get IANA timezone value from 'Country'
    df_filtered[['Local Start Time', 'Timezone']] = df_filtered.apply(apply_conversion_and_capture_timezone, axis=1)
    print("Converted to local time")  # Debug print statement

    # Create local start date to compare to last_record_date to filter
    df_filtered['Local Start Time xTimezone'] = df_filtered['Local Start Time'].astype(str)
    df_filtered['Local Start Time xTimezone'] = df_filtered['Local Start Time xTimezone'].str.slice(stop=-6)
    df_filtered['Local Start Time xTimezone'] = pd.to_datetime(df_filtered['Local Start Time xTimezone'])
    df_filtered['Local Start Date'] = df_filtered['Local Start Time xTimezone'].dt.date
    print("Prepared date comparison")  # Debug print statement

    # Filter by retrieve log last record date
    df_filtered = df_filtered[df_filtered['Local Start Date'] > last_record_date].copy()
    print(f"Filtered by last record date:\n{df_filtered.head()}")  # Debug print statement
    
    # Sort the DataFrame by 'Local Start Time' in ascending order
    df_sorted = df_filtered.sort_values(by='Local Start Date', ascending=True)
    print("Sorted DataFrame")  # Debug print statement
    
    # Save df_filtered to a CSV file
    df_sorted.to_csv(output_file, index=False)
    print(f"Filtered data saved to {output_file}")  # Debug print statement

    print(df_sorted.head())  # Debug print statement

    # Iterate over rows and create calendar events
    for index, row in df_sorted.iterrows():
        try:
            print(f"Creating event for: {row['Title']}")  # Debug print statement
            create_calendar_event(access_token, row)
            print("last_record_date", last_record_date)  # Debug print statement
            print(f"Event created for 'Start Time' {row['Start Time']} 'Local Start Date' {row['Local Start Date']} {row['Title']}")
            #break
        except requests.exceptions.HTTPError as e:
            print(f"HTTPError: {e.response.status_code} - {e.response.text}")  # Debug print statement
            pass
        except Exception as e:
            print(f"General exception: {e}")  # Debug print statement
            pass

if __name__ == "__main__":
    main()
