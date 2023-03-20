# Import necessary functions from utils
from utils import get_outlook_calendar, get_jira_issues, plan_jira_issues, clean_calendar
# Import constants from utils
from utils import EMAIL, JIRA_DOMAIN, JIRA_EMAIL, START_DATE, START_DATE_STR, END_DATE_STR
from colorama import Fore, Style
import datetime

# Main function definition
def main():

    # Initialize outlook calendar
    print(Fore.CYAN + 'Getting data from outlook calendar: ' + EMAIL + Style.RESET_ALL)
    outlook_calendar = get_outlook_calendar()
    
    # Configure outlook items to incude reccurent meetings and sort them by date
    outlook_calendar.Items.IncludeRecurrences = True
    outlook_calendar.Items.Sort("[Start]")

    # Retrieve Jira issues
    print(Fore.CYAN + 'Getting Jira issues from domain ' + JIRA_DOMAIN + " and user " +  JIRA_EMAIL + Style.RESET_ALL)
    issues = get_jira_issues()

    # Get the start and end dates of the current sprint
    sprint_start = START_DATE
    #sprint_end = END_DATE

    # Define the start and end times for working days (working hours)
    workday_start = datetime.time(hour=8)
    #workday_end = datetime.time(hour=17)
    
    print(Fore.CYAN + "Planning sprint from: " + START_DATE_STR + " to " + END_DATE_STR + Style.RESET_ALL)

    # Plan Jira issues
    plan_jira_issues(issues, outlook_calendar, sprint_start, workday_start)
    
    # Clean calendar
    clean_calendar()
                
    print(Fore.GREEN + "PLANNING SUCCESSFUL" + Style.RESET_ALL)
                
                
if __name__ == '__main__':
    main()