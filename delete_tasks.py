# Import necessary functions from utils
from utils import get_outlook_calendar, get_jira_issues, plan_jira_issues
# Import constants from utils
from utils import EMAIL, JIRA_DOMAIN, JIRA_EMAIL, START_DATE, START_DATE_STR, END_DATE_STR

outlook_calendar = get_outlook_calendar()
outlook_calendar_items = outlook_calendar.Items

for item in outlook_calendar_items:
    if item.Categories == "Task":
        item.Delete()
