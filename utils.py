import datetime
import pytz
from jira import JIRA
import win32com.client
from colorama import Fore, Style

# Constants definition
EMAIL = "dorian.popovic@hotmail.com"                                      # Insert outlook e-mail here
FOLDER = "calendar"                                                       # Outlook calendar folder
START_DATE_STR = "2023-03-06 08:00"                                       # Sprint start date
END_DATE_STR = "2023-03-17 17:00"                                         # Sprint end date
START_DATE = datetime.datetime.strptime(START_DATE_STR, '%Y-%m-%d %H:%M') # Datetime convert
END_DATE = datetime.datetime.strptime(END_DATE_STR, '%Y-%m-%d %H:%M')     # Datetime convert
WORK_START = datetime.time(hour=8, minute=0)                              # Working hours start for you
WORK_END = datetime.time(hour=17, minute=0)                               # Working hours end for you
TIMEZONE = pytz.timezone('Europe/Zurich')                                 # Timezone (not sure it is needed)
CATEGORY = "Task"                                                         # jira meetings category (create "Task" one or your own and modify accordingly)
JIRA_DOMAIN = "https://dorianpopovic.atlassian.net/"                      # Insert your Jira workspace/domain here
JIRA_EMAIL = "contact.dpopovic@gmail.com"                                 # Insert Jira email here (if different from outlook email)
JIRA_TOKEN = ""                                                           # Insert Jira generated token here

# Function to retrieve data from the outlook calendar
def get_outlook_calendar():
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        outlook_calendar = outlook.Folders[EMAIL].Folders[FOLDER] # Retrieve your calendar data
        
        return outlook_calendar
    
# Function to retrieve data from Jira issues
def get_jira_issues():
    jira = JIRA(JIRA_DOMAIN, basic_auth=(JIRA_EMAIL, JIRA_TOKEN)) # Initialize Jira instance
    jql = f'assignee = currentUser()' # Get Jira issues assigned to you in the current sprint
    issues = jira.search_issues(jql)
    
    return issues

def plan_jira_issues(issues, outlook_calendar, sprint_start, workday_start):
    
    # Loop through each issue and schedule working hours in the calendar
    for issue in issues:
        
        # Get the estimated time for this issue
        story_points = issue.fields.customfield_10016
        estimated_time = story_points * 60  # Convert story points to minutes (easier operations)
        
        # Loop through each day of the sprint
        for day in range(7):
            date = sprint_start + datetime.timedelta(days=day)
            
            # Skip weekends
            if date.weekday() >= 5:
                continue
            
            # Loop through each hour of the workday
            for hour in range(9):
                start_time = datetime.datetime.combine(date, workday_start) + datetime.timedelta(hours=hour)
                end_time = start_time + datetime.timedelta(minutes=60)
                
                start_time_str = start_time.strftime("%Y-%m-%d %H:%M %p")
                end_time_str = end_time.strftime("%Y-%m-%d %H:%M %p")
                
                # Check if this hour is available
                appointments = outlook_calendar.Items.Restrict("[Start] <= '" + start_time_str + "' AND [End] >= '" + end_time_str + "'")
                
                found_appointment = False # Initiliaze to false, weird way of .Restrict to collects items
                
                for appointment in appointments:
                    if appointment.Start <= start_time.replace(tzinfo=pytz.UTC) and appointment.End >= end_time.replace(tzinfo=pytz.UTC):
                        found_appointment = True
                        break
                    
                if not found_appointment:
                    
                    # No appointment found so we can schedule this hour for the current issue
                    appointment = outlook_calendar.Items.Add()
                    appointment.Subject = issue.fields.summary
                    appointment.Start = start_time.replace(tzinfo=pytz.UTC)
                    appointment.End = start_time.replace(tzinfo=pytz.UTC) + datetime.timedelta(minutes=60)
                    appointment.Categories = CATEGORY
                    appointment.BusyStatus = 1 # Tentative
                    appointment.Save()
                    
                    # Subtract this hour from the estimated time
                    estimated_time -= 60
                    
                    # If all time has been scheduled, move on to the next issue
                    if estimated_time == 0:
                        break
                        
            # If all time has been scheduled, move on to the next issue
            if estimated_time == 0:
                break

# Function to clean calendar and merge consecutively planned meetings                
def clean_calendar():
    
    outlook_calendar = get_outlook_calendar()
    outlook_calendar_items = outlook_calendar.Items
    previous_item = None
    items_to_delete = []
    start_time = None
    end_time = None
    meeting_to_update = None

    for item in outlook_calendar_items:
        if item.Categories == "Task":
            if previous_item!= None:
                if item.Subject == previous_item.Subject and item.Start == previous_item.End:
                    items_to_delete.append(item)
                    if start_time is None:
                        # This is the first duplicated task, so set the start time to the start time of this task
                        start_time = previous_item.Start
                        meeting_to_update = previous_item
                    # Update the end time to the end time of the current duplicated task
                    end_time = item.End

                    if start_time is not None and end_time is not None:
                        meeting_to_update.Start = start_time
                        meeting_to_update.End = end_time
                        meeting_to_update.Save()

                else:
                    # This is not a duplicated task, so reset the start and end times
                    start_time = None
                    end_time = None
                    meeting_to_update = None
            previous_item = item
        else:
            continue

    for item in items_to_delete:
        item.Delete()
    
    print(Fore.MAGENTA + "Cleaned calendar" + Style.RESET_ALL)
