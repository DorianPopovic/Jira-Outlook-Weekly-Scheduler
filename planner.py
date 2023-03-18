# Import necessary functions or modules
from utils import *
import datetime
import pytz
from jira import JIRA
import win32com.client
from colorama import init, Fore, Style

def main():

    # Initialize outlook calendar
    print(Fore.CYAN + 'Getting data from outlook calendar: ' + EMAIL + Style.RESET_ALL)
    outlook_calendar = get_outlook_calendar()

    outlook_calendar.Items.IncludeRecurrences = True
    outlook_calendar.Items.Sort("[Start]")

    # Retrieve Jira issues
    print(Fore.CYAN + 'Getting Jira issues from domain ' + JIRA_DOMAIN + " and user " +  JIRA_EMAIL + Style.RESET_ALL)
    issues = get_jira_issues()

    # Get the start and end dates of the current sprint
    sprint_start = START_DATE
    sprint_end = END_DATE

    # Define the start and end times for working hours
    workday_start = datetime.time(hour=8)
    workday_end = datetime.time(hour=17)
    
    print(Fore.CYAN + "Planning sprint from: " + START_DATE_STR + " to " + END_DATE_STR + Style.RESET_ALL)


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
                    appointment.BusyStatus = 2 # Tentative
                    appointment.Save()
                    
                    # Subtract this hour from the estimated time
                    estimated_time -= 60
                    
                    # If all time has been scheduled, move on to the next issue
                    if estimated_time == 0:
                        break
                        
            # If all time has been scheduled, move on to the next issue
            if estimated_time == 0:
                break
                
    print(Fore.GREEN + "PLANNING SUCCESSFUL" + Style.RESET_ALL)
                
                
if __name__ == '__main__':
    main()