# Google Calendar Event Extractor

This Python script extracts events from Google Calendar using the Google Calendar API and generates a CSV report.

## Key Features

- **Data Extraction:** Retrieves events from Google Calendar using the Google Calendar API.
- **Time Zone Conversion:**: Converts event times to the local time zone.
- **Data Processing:** Parses event details and calculates event durations.
- **Data Filtering:**Filters events based on user input for start and end dates.
- **Export to CSV:** Generates a CSV report with event details.
- **Report Generation:**Optionally generates reports for individual meetings based on extracted events.

## How to Use

### Prerequisits

1. **Python 3**: Ensure you have Python 3 installed on your system. You can check the Python version by running python --version in the command line. 
2. **Dependencies**: 
- `google-api-python-client` library
- pytz library
- Google Calendar API credentials
- Google Calendar IDs
- API key

### Installation

1. Install the required libraries:
   ```bash
   pip install google-api-python-client
   pip install pytz

2. Obtain Google Calendar API credentials and an API key by following the instructions in the Google Calendar API documentation. [Google Calendar API documentation](https://support.google.com/googleapi/answer/6158862?hl=en)

3. Obtain Google Calendar ID by going to Google Calendar and signing in, if you're not already signed in. In the upper menu bar, click on the gear icon (⚙️) to open the Settings menu, then select "Settings". In the Settings menu, locate the calendar for which you want to find the ID under the "Settings for my calendars" section. Click on the name of the calendar to view its settings. Scroll down to the "Integrate calendar" section. Your Calendar ID is displayed under "Calendar ID".

4.  Ensure Google Calendar is available to the public. In the left menu bar, click on "Calendar settings" for the desired calendar. Scroll down to the "Access permissions" section. Check the box next to "Make available to public" to enable this option.

5.    Replace the placeholders for api_key and calendar_id with your actual API key and calendar IDs.

### Running the Script

1. **Clone Repository**: Clone the repository containing the script.
2. **Navigate to Directory**: Open terminal and navigate to the directory containing the script.
3. **Run Accounting Helper Script**: Execute the script by running:
    ```bash
    python3 Google_Cal_Meeting_Export.py
    ```
4. **Follow Instructions**: Enter the start date and end date for the meetings you would like to export.

5. **Exported Data**: Once processed, the modified data will be exported to a CSV file named "Meetings_{start_date}_to_{end_date}.csv".

6. **Optional Step**: Generate a report that will print all meetings with a specific event title to the screen between your designated start date and end date, along with the total number of hours spent on that event title. This feature is useful if you have recurring meetings with the same title and want to track the total time spent on it. For example, if you have weekly meetings titled "Planning", you can use this option to determine how many hours per week you spent on planning. The program will print the details of each meeting, including the date and duration, as well as the total hours spent on the specified event title within the designated date range. For best results, title your events consistently for easy tracking. 

## Known Issues

None reported.
