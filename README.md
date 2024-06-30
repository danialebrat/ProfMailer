# ProfMailer

**Automate the Process of Emailing University Professors for Graduate Positions**

ProfMailer is a desktop application designed to help you efficiently send personalized emails to university professors to inquire about graduate positions. You can run it as a standalone application or modify the open-source files to create your own customized version.

## Requirements

- **Gmail Account**
- **Set Up 2-Step Verification**
  - [Google Account Security](https://myaccount.google.com/security)
- **Create an App Password for the Application**
  - [App Passwords](https://security.google.com/settings/security/apppasswords)
- **Email Templates**
- **Resumes**
- **Transcript**
- **Completed CSV File** (based on sample file)

## How to Use

To use ProfMailer, prepare the following files:

### CSV File

Create a CSV file containing professor information. Refer to the `list.csv` in the example files. The CSV should include the following columns:

- **name**: Last name of the recipient (without the title "Prof.")
- **email**: Email address of the recipient
- **university**: Name of the university
- **send_status**: Initially set to 0; the program will update it to 1 after sending an email to avoid duplicates
- **subject**: Subject of your email (use "Prospective Graduate Student" if unspecified)
- **transcript**: True if the professor needs it; otherwise, False
- **field**: Research field of the professor (each field should have a corresponding resume)
- **group**: Lab name, if applicable; otherwise, "none"

### Folder Structure

Create a folder containing:

- All resumes
- Email templates
- Transcript

### Gmail Account

- **Email**: e.g., sample@gmail.com
- **App Password**: A Google app password (NOT your Gmail password)
  - Example: `abcd efgh ijkl mnop`

### Resume Prefix

- For each **field**, create a corresponding **PDF** résumé.
- The resume prefix should be followed by `_{field}`.
  - Example: If the prefix is `Danial_Ebrat_CV` and you are applying for AI and Music, the resumes should be named `Danial_Ebrat_CV_AI.pdf` and `Danial_Ebrat_CV_Music.pdf`.
  - Ensure the **field** values in the CSV match the filenames.

### Email Template

- For each **field**, create a corresponding **HTML** template.
- Template filenames should follow the format `template_{field}`.
  - Example: `template_AI.html` and `template_Music.html`
- Use tools like [Word to HTML](https://wordhtml.com/) if you are unfamiliar with HTML.

### Transcript

- Have your transcript file as a PDF, named `transcript.pdf`.

### Running ProfMailer

1. Specify the path to the CSV file.
2. Specify the path to the folder containing all files (resumes, transcript, templates).
3. Enter your email and app password.
4. Enter the resume prefix (e.g., `Danial_Ebrat_CV`).
5. Click "Send Emails" and let ProfMailer do the work.

## Important Notes

- **Python Developers**: Instead of running `ProfMailer.exe`, clone the project and modify `ProfMailer.py` as needed.
- The program randomly waits 1-2 minutes between each email to comply with Google’s email sending limits.
- **VERY IMPORTANT**: Do not send more than 30 emails per hour to avoid being banned, flagged, or marked as spam.

