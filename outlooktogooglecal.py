import win32com.client
import time
import datetime
import config



outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
appointments = namespace.GetDefaultFolder(9).Items

appointments.Sort("[Start]")
appointments.IncludeRecurrences = "True"

# start date
now = datetime.datetime.now()
monthNow = now.month
yearNow = now.year
begin = datetime.date(now.year, now.month, 1)
# end date
nextMonth = begin + datetime.timedelta(days=35)
end = datetime.date(nextMonth.year, nextMonth.month, 1)
print("[Start] >= '" + begin.strftime("%m/%d/%Y") +   "' AND [End] '" + end.strftime("%m/%d/%Y")+ "'")
appointments = appointments.Restrict("[Start] >= '" + begin.strftime("%m/%d/%Y") +   "' AND [End] <='" + end.strftime("%m/%d/%Y")+ "'")


print appointments
for appt in appointments:
    print ("Subject: {0} --  Start : {1} -- End: {2}".format(appt.Subject, appt.Start, appt.End))


##format #date MM/DD/YYYY format.
#header =
#'“Subject”,
#“Start Date”, = MM/DD/YYYY
#“Start Time”, = 24 hour time (13:45) or 12 AM/PM (01:45 PM)
#“End Date”, = MM/DD/YYYY
# “End Time”, = 24 hour time (13:45) or 12 AM/PM (01:45 PM)
# “All Day Event”, = True / False
#  “Description”,
#  “Location”,
 #“Private”' True / False
header = '“Subject”, “Start Date”, “Start Time”, “End Date”, “End Time”, “All Day Event”, “Description”, “Location”, “Private”'
