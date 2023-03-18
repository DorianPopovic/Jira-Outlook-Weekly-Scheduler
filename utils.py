import datetime
import pytz
from jira import JIRA
import win32com.client

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
CATEGORY = "Task"                                                         # Color meeting category
JIRA_DOMAIN = "https://dorianpopovic.atlassian.net/"                      # Insert your Jira workspace/domain here
JIRA_EMAIL = "contact.dpopovic@gmail.com"                                 # Insert Jira email here (if different from outlook email)
JIRA_TOKEN = ""

# Function to retrieve data from the outlook calendar
def get_outlook_calendar():
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        outlook_calendar = outlook.Folders[EMAIL].Folders[FOLDER] # Retrieve your calendar data
        #outlook_calendar_items = outlook_calendar_folder.Items
        #outlook_calendar_items.IncludeRecurrences = True
        #outlook_calendar_items.Sort("[Start]")
        #outlook_calendar_items = outlook_calendar_items.Restrict("[Start] >= '" + START_DATE_STR + "' AND [END] <= '" + END_DATE_STR + "'") # Filter calendar for current sprint
        
        return outlook_calendar
    
# Function to retrieve data from Jira issues
def get_jira_issues():
    jira = JIRA(JIRA_DOMAIN, basic_auth=(JIRA_EMAIL, JIRA_TOKEN)) # Initialize Jira instance
    jql = f'assignee = currentUser()' # Get Jira issues assigned to you in the current sprint
    issues = jira.search_issues(jql)
    
    return issues
