

ROSTER_DIR = "."
ROSTER_FILE = f"{ ROSTER_DIR }/staff_roster.xls"

# origin 1 row of the title line in the roster file
ROSTER_TITLE_ROW = 6

#VC_DR_ID = '1849'     # DR716-21 TX Winter Storm
#VC_DR_ID = '1881'     # DR767-21 Border Ops
VC_DR_ID = '2034'   # DR155-21 Gold Country Forest Fires
DR_NAME = 'DR155-22'

OUTPUT_DIR = "."
OUTPUT_FILE = f"{ OUTPUT_DIR }/staffing.xlsx"

OUTPUT_SHEET_REPORTING = "Reporting"
OUTPUT_SHEET_NOSUPS = "NoSupervisors"
OUTPUT_SHEET_NONSVS = "NoReports"
OUTPUT_SHEET_3DAYS = "Days3Left"


WORKFORCE_SITE_ID = 'americanredcross.sharepoint.com,961a64a8-a0ec-4bd8-b882-0e2351401602,2dc8cace-9fc1-4e91-b3e2-04eb8f57ccab'
WORKFORCE_DRIVE_ID = 'b!qGQaluyg2Eu4gg4jUUAWAs7KyC3Bn5FOs-IE649XzKthyVvQHGwASKQvZvMgibGh'
WORKFORCE_FOLDER_PATH = '/Workforce/.Archive/DO NOT DELETE - AUTO EMAIL DATA/Mail Merge Spreadsheets'

#MAIL_OWNER = "DR767-21-Staffing-Reports-Owner@AmericanRedCross.onmicrosoft.com"
#MAIL_ADDRESS = 'dr767-21-staffing-reports@americanredcross.onmicrosoft.com'
#MAIL_ARCHIVE = 'https://outlook.office.com/mail/group/americanredcross.onmicrosoft.com/dr767-21-staffing-reports/email'

#MAIL_OWNER = "neil.katin@redcross.org"
MAIL_OWNER = "dr155-22Staffing@redcross.org"
MAIL_SENDER = MAIL_OWNER
MAIL_ADDRESS = 'dr155-22-staffing-reports@americanredcross.onmicrosoft.com'
MAIL_BCC = 'dr155-22-test-messages@askneil.com'
MAIL_ARCHIVE = 'https://outlook.office.com/mail/group/americanredcross.onmicrosoft.com/dr155-22-staffing-reports/email'
DAYS_BEFORE_WARNING = 4

COOKIE_FILE = 'cookies.txt'
WEB_TIMEOUT = 60
