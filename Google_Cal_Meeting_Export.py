import csv
from googleapiclient.discovery import build
from datetime import datetime, timedelta
import pytz
import math
from dotenv import load_dotenv
import os

# Load variables from the .env file
load_dotenv()

# Access the variables
SECRET_API_KEY = os.getenv("SECRET_API_KEY")
SECRET_CALENDAR_ID = os.getenv("SECRET_CALENDAR_ID")
DEBUG = os.getenv("DEBUG")

# Use environment variables for API key and calendar ID
api_key = SECRET_API_KEY
calendar_id = SECRET_CALENDAR_ID

# Set your calendar timezone
calendar_timezone = "America/Los_Angeles"

# Prompt the user to input the start date and end date
start_date_str = input("Enter the start date (YYYY-MM-DD): ")
end_date_str = input("Enter the end date (YYYY-MM-DD): ")

# Parse start and end dates into datetime objects
start_date_obj = datetime.strptime(start_date_str, "%Y-%m-%d").date()
end_date_obj = datetime.strptime(end_date_str, "%Y-%m-%d").date()

# Retrieve event list using Google Calendar API.
service = build("calendar", "v3", developerKey=api_key)
events_result = service.events().list(
    calendarId=calendar_id,
    timeMin=start_date_str + "T00:00:00Z",  # Append time component for start date
    timeMax=end_date_str + "T23:59:59Z",  # Append time component for end date
    fields="items(summary,start,end)",
    timeZone="UTC"
).execute()
 
# Prepare the lists to store event data
sessions = [["Event title", "Date", "Start Time", "End Time", "Event Duration", "Converted Event Duration"]]
all_day_events = [["Event title", "Date", "Start Time", "End Time", "Event Duration", "Converted Event Duration"]]

# Define the Pacific Time zone
timezone = pytz.timezone(calendar_timezone)

# Process events and store data
for event in events_result.get("items", []):
    # Extract start datetime and date
    start_datetime = event.get("start", {}).get("dateTime")

    # If start_datetime is provided, convert it to local timezone
    if start_datetime:
        start_time = datetime.fromisoformat(start_datetime)
        start_time = start_time.astimezone(timezone)

        # Check if event occurs on or after the specified start_date
        if start_time.date() >= start_date_obj:
            # Extract end datetime
            end_datetime = event.get("end", {}).get("dateTime")

            # If end_datetime is provided, convert it to local timezone
            if end_datetime:
                end_time = datetime.fromisoformat(end_datetime)
                end_time = end_time.astimezone(timezone)

                # Calculate the event duration
                duration = end_time - start_time

                # Convert duration to hours and round up to the nearest 0.25
                duration_hours = duration.total_seconds() / 3600
                duration_rounded = math.ceil(duration_hours * 4) / 4

                # Format duration with two decimal places
                formatted_duration = "{:.2f}".format(duration_rounded)

                # Extract event title
                event_title = event.get("summary", "")

                # Remove "Tutoring" from the event title and trim any leading/trailing spaces
                event_name = event_title.split(" Tutoring")[0].strip()

                # Append event details to the sessions list
                sessions.append([
                    event_name,
                    start_time.strftime("%m/%d/%Y"),  # Format start date as MM/DD/YYYY
                    start_time.strftime("%I:%M %p"),  # Format start time as HH:MM AM/PM
                    end_time.strftime("%I:%M %p"),  # Format end time as HH:MM AM/PM
                    str(duration),  # Event duration in hours and minutes
                    formatted_duration  # Converted event duration
                ])
    else:  # Process all-day events
        # Extract event details
        event_title = event.get("summary", "")
        event_date = event.get("start", {}).get("date", "")

        # Check if event_date is not 'Date'
        if event_date != 'Date':
            # Set start time to midnight
            start_time = datetime.strptime(event_date, "%Y-%m-%d").replace(hour=0, minute=0, second=0)
            start_time = start_time.astimezone(timezone)

            # Set end time to end of day (23:59:59)
            end_time = datetime.strptime(event_date, "%Y-%m-%d").replace(hour=23, minute=59, second=59)
            end_time = end_time.astimezone(timezone)

            # Calculate event duration for all-day events (set to 0.00)
            event_duration = "0.00"

            # Set converted event duration to 0.00 for all-day events
            converted_event_duration = "0.00"

            # Append event details to all_day_events
            all_day_events.append([
                event_title,
                start_time.strftime("%m/%d/%Y"),  # Format start date as MM/DD/YYYY
                start_time.strftime("%I:%M %p"),  # Format start time as HH:MM AM/PM
                end_time.strftime("%I:%M %p"),  # Format end time as HH:MM AM/PM
                event_duration,
                converted_event_duration
            ])

# Prepare the lists to store event data
events_data = []
sessions_data = []
all_day_events_data = []
events_header = ["Event title", "Date", "Start Time", "End Time", "Event Duration", "Converted Event Duration"]

# Append header row for sessions
sessions_data.append(events_header)

# Append sessions without the header row
sessions_data.extend(sessions[1:])

# Append all_day_events without the header row
all_day_events_data.extend(all_day_events[1:])

# Combine both sessions_data and all_day_events_data into a single list
events_data.extend(sessions_data)
events_data.extend(all_day_events_data)

# Sort events_data by event title and then by date, excluding rows with 'Date'
events_data_sorted = sorted(events_data[1:], key=lambda x: (x[0], datetime.strptime(x[1], "%m/%d/%Y")) if x[1] != 'Date' else ('',))

# Output the retrieved values as a CSV file
output_file = f"/Users/altuser/Documents/Tutoring/Meetings_{start_date_str}_to_{end_date_str}.csv"
with open(output_file, "w", newline="") as f:
    writer = csv.writer(f)
    writer.writerow(events_data[0])  # Write header row
    writer.writerows(events_data_sorted)  # Write sorted events data to CSV

print("Export was successfully written to:", output_file)


######PART 2###########
while True:

    def calculate_student_hours(sessions, student_name, start_date, end_date):
        total_hours = 0
        for session in sessions:
            event_title = session[0].strip()  # Get event title
            if student_name in event_title:
                start_date_str = session[1]  # Get start date string
                end_date_str = session[1]  # Get end date string

                # Parse start and end dates into datetime objects
                start_time = datetime.strptime(start_date_str, "%m/%d/%Y").date()
                end_time = datetime.strptime(end_date_str, "%m/%d/%Y").date()

                if start_time >= start_date and end_time <= end_date:
                    duration_rounded = float(session[5])  # Extract converted event duration
                    total_hours += duration_rounded
        return total_hours

    # Create a list to store student names
    student_names = []

    # Iterate through filtered sessions and extract student names from event titles
    for session in sessions:
        event_title = session[0].strip()  # Strip leading and trailing spaces
        # Check if "Tutoring" exists in the event title
        if "Tutoring" in event_title:
            # Remove " Tutoring" from the event title and trim any leading/trailing spaces
            name = event_title.split(" Tutoring")[0].strip()
        else:
            # If "Tutoring" doesn't exist, use the event title directly
            name = event_title.strip()
        # Append the student name to the list
        student_names.append(name)

    # Get unique student names
    unique_student_names = list(set(student_names))

    # Filter out "Event Title" if it exists
    if "Event title" in unique_student_names:
        unique_student_names.remove("Event title")

    # Create a list to store student data
    student_data = []

    # Ask the user if they want to generate a report
    generate_report = input("Would you like to generate a report? (y/n): ").lower()
    if generate_report == "y":
        print("Which student(s) would you like to report on?")
        for index, name in enumerate(unique_student_names, start=1):
            print(f"{index}. {name}")

        student_choice = int(input("Enter the number of the student: "))

        # Filter sessions based on the selected student
        if 0 < student_choice <= len(unique_student_names):
            selected_name = unique_student_names[student_choice - 1]
            total_hours = calculate_student_hours(sessions_data, selected_name, start_date_obj, end_date_obj)

            # Sort sessions for the selected student by start date before printing
            sorted_sessions = sorted((session for session in sessions[1:] if selected_name in session[0]), key=lambda x: datetime.strptime(x[1], "%m/%d/%Y"))

            # Loop through sorted sessions for the selected student
            for session in sorted_sessions:
                event_title = session[0].strip()  # Strip leading and trailing spaces
                clean_event_title = event_title.split(" Tutoring")[0].strip()
                start_date = datetime.strptime(session[1], "%m/%d/%Y")  # Parse start date
                duration_rounded = float(session[5])  # Extract converted event duration
                print(f"{clean_event_title.strip()} {start_date.strftime('%m/%d/%Y')} {duration_rounded:.2f} hours")


            # Print the total hours for the selected student
            print(f"{selected_name} had a total of {total_hours} hours between {start_date_str} and {end_date_str}")

            # Store student data
            student_data.append([selected_name, total_hours])
        else:
            print("Invalid input. Exiting.")
    elif generate_report == "n":
            print("Thank you. Goodbye.")
            break
else:
    print("Thank you. Goodbye.")
