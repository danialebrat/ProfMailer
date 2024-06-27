# ProfMailer
Automatically send emails to university professors to find Graduate positions
ProfMailer can be used as a desktop application to run, or you can modify the open-source files and create your own version. 
## Requirements:

- Gmail account
- set up 2-Step Verification 
  - Link: https://myaccount.google.com/security
- Creating an App password for the application
  - link: https://security.google.com/settings/security/apppasswords


## How to use it
To use this program, you need to have below file (see sample files):

- a CSV file containing professors information: see list.csv in example file
  - **name**: Last name of the receiver without the title (the program will add "Prof." in the beginning)
  - **email**: email address of the receiver
  - **university**: name of the university
  - **send_status**: initially it should be 0, the program will change it to 1 after sending an email to avoid sending duplicates
  - **subject**: subject of your email; some professors particularly define the subject (e.g. PhD Fall 2023). If the subject is "none", the program will use "Prospective Graduate Student" 
  - **transcript**: True if the professor needs it, otherwise, False
  - **field**: research field of the professor. for each field, you should have a separate resume for applying to that field. This will give you the power to specify different resumes to send for different professors.
  - **group**: if the professor has a Lab name, put it in this column, otherwise, enter "none".
- Resume:
  - for every **field**, you should have a **pdf** resume. 


- Download ProfMailer.exe 
- run
- Those who knows python, ju