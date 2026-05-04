Requirements:

Keep your work mail as the defeault mail on outlook.
File > Options > Mail > Send Mail > New drafts from default mail

Save Mail Template as Oft.
In the body keep first Line as : Dear [First name],
Also keep [sender name] and [sender mail] wherever you want them to be replaced.

In your database ensure Names column is filled with names or placeholders like []. The program ends once blank name is encountered.

Enter column names when prompted.

Change paths for database csv, Mail Template oft and set mail address.

Ensure python is installed.
Run in terminal:
python -m pip install --upgrade pip
python -m pip install pywin32 openpyxl

Sign in with all emails on outlook classic for pc
Then Restart the pc

(Optional : Create a folder in your main mail for the project and then create a rule from settings such that all mails with certain keywords from responses are routed to that folder)

To start mail automation Run from Mailer folder: python main.py
Enter mail ids seperated by comma (,)
eg: 
richard.m@ureka.co.uk,mia.anderson@ureka.co.uk,neha.agarwal@ureka.co.uk,priya.iyer@ureka.co.uk,meena.ghosh@ureka.co.uk

Change the path on the main file with your mail template and database paths
Mode 1: 
TEMPLATE_PATH = r"C:\Users\xenor\AppData\Roaming\Microsoft\Templates\Ureka_multi_aifsi.oft"
DATABASE_PATH = r"C:\AutoEdge\New folder\Multi_Mailer\AIfSI KPI Dashboard.xlsx - Richard.csv"


Mode 2: ENROLL
TEMPLATE_PATH = r"C:\Users\xenor\AppData\Roaming\Microsoft\Templates\Ureka_enroll.oft"
DATABASE_PATH = r"C:\AutoEdge\New folder\Multi_Mailer\Enrollment KPI Tracker.xlsx - Sheet4.csv"

Mode 3: LAUNCH
TEMPLATE_PATH = r"C:\Users\xenor\AppData\Roaming\Microsoft\Templates\Ureka_Launch.oft"
DATABASE_PATH = r"C:\Users\xenor\Documents\Multi_Mailer\Active Enrollment - Sheet1.csv"