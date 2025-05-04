# Email Feedback App
A desktop application for processing positive feedback from Excel files and generating email drafts in Microsoft Outlook. Designed for service desk teams to recognize analysts based on customer satisfaction feedback.
## ✨ Features
- **Import Feedback from Excel Files**: Load feedback data from Excel files per account, stored in the **data/** folder.
- **Filter Valid Feedbacks**: Apply custom filtering rules (e.g., rating, assignment groups) to identify valid feedbacks.
- **Manual Editing**: Edit feedback messages and analyst names directly in the interface via double-click.
- **Approve or Reject Feedbacks**: Approve feedbacks to generate email drafts or reject them to remove from the list.
- **Prevent Duplicates**: Use a log system **(logs/approved_feedbacks.xlsx)** to prevent processing the same feedback multiple times.
- **Generate Email Drafts**: Create visually styled email drafts in Outlook (not sent automatically) with analyst info, feedback message, and embedded images.
- **Missing Email Warning**: Display a ⚠️ symbol next to analysts with missing email addresses in the **analysts.json** configuration.
- **Paginated Interface**: Manage large numbers of feedbacks with pagination (20 items per page).
- **Export Approved Feedbacks**: Log approved and rejected feedbacks to a dedicated Excel file **(logs/approved_feedbacks.xlsx)**.
- **Account Selection**: Select accounts from a dropdown menu to load and process feedbacks specific to that account.
- **Settings Window**: Configure settings such as default sender email and template language (Portuguese, English, Spanish) via a settings window.
- **Refresh Data**: Reload feedback data from Excel files without restarting the application using the "Refresh Data" button.
- **Visual Feedback**: Display a loading cursor during time-consuming operations (e.g., loading feedbacks, generating emails).
## 📁 Folder Structure
```text
email_feedback_app-main/
│
├── main.py                         # App launcher
├── config/
│   ├── config.json                # Account-specific Excel mapping and app settings
│   └── analysts.json              # Analyst names and emails per account/group
├── data/
│   └── [excel files here]         # Raw feedback Excel files (e.g., Flowserve.xlsx)
├── logs/
│   ├── approved_feedbacks.xlsx    # Log of approved/rejected feedbacks (prevents duplicates)
│   └── error.log                  # (Optional) Log file for errors (if implemented)
├── templates/
│   ├── email_template.html        # Email HTML template (Portuguese)
│   ├── email_template_en.html     # Email HTML template (English)
│   ├── email_template_es.html     # Email HTML template (Spanish)
│   └── assets/
│       ├── header.png             # Header image for email templates
│       └── Award-Winner.png       # Award image for email templates
├── email_feedback_app/
│   ├── ui.py                      # GUI interface logic
│   ├── processor.py               # Excel processing logic
│   ├── utils.py                   # Helper functions (email generation, log handling, etc.)
│   ├── settings_window.py         # Settings window logic
│   └── __init__.py                # Package initialization```
```
## 🛠️ Requirements
- **Python 3.11 or higher**
- **Microsoft Outlook** installed (required for generating email drafts)
- **Windows OS** (due to **win32com.client** dependency for Outlook integration)
### Required Python Packages:
- **openpyxl**: For reading and writing Excel files.
- **pywin32**: For Outlook integration via **win32com.client**.
### Install Dependencies:
```bash
pip install openpyxl pywin32
```
## 🚀 How to Run
#### 1. Clone or Download the Repository:
 - Download the project folder **email_feedback_app-main** or clone it using Git:
```bash
git clone <repository-url>
cd email_feedback_app-main
```
#### 2. Set Up the Environment:
- Ensure Python 3.11+ is installed.
- Install the required packages (see above).
- Ensure Microsoft Outlook is installed and configured on your system.
#### 3. Prepare the Data:
- Place your feedback Excel files in the **data/** folder. Each file should be named after the account it represents (e.g., **Flowserve.xlsx** for the "Flowserve" account).
- Configure the **config/config.json** file with the correct column mappings for each account (see "How to Add a New Account" below).
- Configure the **config/analysts.json** file with analyst names and email addresses (see "Analysts Configuration" below).
#### 4. Run the Application:
- Open a terminal in the project directory and run:
```bash
python main.py
```
#### 5. Using the Application:
- The application window will open with the title "Kudos Manager".
- Select an account from the dropdown menu.
- Click "Load Feedbacks" to display the feedbacks for the selected account.
- Double-click on a feedback's "Message" or "Analyst" column to edit the content.
- Click the "Delete" button in the "Action" column (or press the **Delete** key) to reject a feedback.
- Select feedbacks in the list and click "Generate Emails" to create email drafts in Outlook.
- Click "Refresh Data" to reload feedback data from the **data/** folder if the Excel files are updated.
- Click "⚙️ Settings" to configure the default sender email and email template language (Portuguese, English, or Spanish).

## 🧩 How to Add a New Account
To add a new account, you need to update the **config/config.json** file with the account's Excel column mappings and filtering rules.

#### Steps:
1. Open **config/config.json** in a text editor.
2. Add a new entry for your account. The key should match the name of the Excel file in the **data/** folder (e.g., **Flowserve** for **Flowserve.xlsx**).
#### Example Entry:
```json
{
  "Flowserve": {
    "sheet_name": "Sheet1",
    "header_row": 1,
    "ticket_id": "A",
    "message": "B",
    "analyst_name": "C",
    "user_name_parts": ["D", "E"],
    "rating": "F"
  }
}
```
#### - Required Fields:
- **sheet_name**: The name of the Excel sheet containing the feedback data (e.g., **"Sheet1"**).
- **header_row**: The row number (1-based) containing the column headers (e.g., **1**).
- **ticket_id**: The column letter for the ticket ID (e.g., **"A"**).
**message**: The column letter for the feedback message (e.g., **"B"**).
- **analyst_name**: The column letter for the analyst's name (e.g., **"C"**).
- **user_name** or **user_name_parts**: Either a single column letter for the user's name (e.g., **"user_name": "D"**) or a list of columns to combine (e.g., **"user_name_parts": ["D", "E"]**).
#### - Optional Filters:
- **rating**: A column letter for a numeric rating (e.g., 1–5). Only feedbacks with a rating of 4 or 5 are considered valid.
```json
"rating": "F"
```
- **rating_text**: Filter based on a text value in a column (e.g., "Satisfied").
```json
"rating_text": {
  "column": "F",
  "positive_value": "Satisfied"
}
```
- **rating_inverted**: Filter based on specific numeric values.
```json
"rating_inverted": {
  "column": "F",
  "valid_values": ["4", "5"]
}
```
- **assignment_group**: Filter feedbacks by assignment group.
```json
"assignment_group": {
  "column": "C",
  "required_value": ["Group A", "Group B"]
}
```
3. Save the **config.json** file.
4. Place the corresponding Excel file (e.g., **Flowserve.xlsx**) in the **data/** folder.
5. Restart the application or click "Refresh Data" to load the new account.

## 👥 Analysts Configuration
The **config/analysts.json** file maps analyst names to their email addresses, organized by account and group.

#### Steps:
1. Open **config/analysts.json** in a text editor.
2. Add an entry for your account, with groups and analysts.
#### Example Entry:
```json
{
  "Flowserve": {
    "groups": {
      "Service Desk-LAC-SAO-FLS": {
        "analysts": {
          "Jane Doe": "jane.doe@example.com",
          "John Smith": "john.smith@example.com"
        },
        "cc_emails": ["manager@example.com"]
      }
    }
  }
}
```
#### - Structure:
- The top-level key is the account name (e.g., **"Flowserve"**).
- **groups**: A dictionary of groups within the account (e.g., **"Service Desk-LAC-SAO-FLS"**).
- **analysts**: A dictionary mapping analyst names to their email addresses.
- **cc_emails**: (Optional) A list of email addresses to be added to the CC field of the email drafts.
3. Save the **analysts.json** file.
4. Restart the application or click "Refresh Data" to apply the changes.
```
⚠️ Note: If an analyst does not have an email address in analysts.json, a ⚠️ symbol will appear next to their name in the interface. The email draft will still be generated, but the "To" field will be empty.
```

## 📨 Email Drafts
Approved feedbacks can be used to generate Outlook email drafts (not sent automatically). The email includes:

- A visually styled HTML message using one of the templates in the **templates/** folder (**email_template.html** for Portuguese, **email_template_en.html** for English, **email_template_es.html** for Spanish).
- The analyst’s name, user name, ticket ID, and feedback message.
- Embedded images (e.g., **header.png**, **Award-Winner.png**) from the **templates/assets/** folder.
- A subject line based on the account and template language (e.g., **[Flowserve] Reconhecimento de Excelente Atendimento** for Portuguese).
### Steps to Generate Email Drafts:
1. Select an account from the dropdown and click "Load Feedbacks".
2. (Optional) Edit the feedback message or analyst name by double-clicking the respective cell in the table.
3. Select the feedbacks you want to process (currently, all displayed feedbacks are processed; see "Future Ideas" for selection improvements).
4. Click "Generate Emails".
5. Confirm the action if prompted (if confirmation is implemented).
6. Check Microsoft Outlook for the generated email drafts in the "Drafts" folder.
## ⚙️ Settings Configuration
The application includes a settings window to configure global options.

#### Accessing the Settings:
- Click the "⚙️ Settings" button in the top-right corner of the main window.
#### Available Settings:
- **Default Sender Email**: The email address used in the "Sent On Behalf Of" field for Outlook drafts.
- **Template Language**: Choose the language for email templates:
  - Portuguese (**email_template.html**)
  - English (**email_template_en.html**)
  - Spanish (**email_template_es.html**)
### Saving Settings:
- After making changes, click "Save" in the settings window.
- Some changes (e.g., template language) may require restarting the application to take effect.
## 📋 Logging & Duplicates
Approved and rejected feedbacks are logged in:

```text
logs/approved_feedbacks.xlsx
```
#### Log Structure:
- **Timestamp**: Date and time of the action.
- **Account**: The account name (e.g., "Flowserve").
- **TicketID**: The ticket ID of the feedback.
- **UserName**: The name of the user who provided the feedback.
- **AnalystName**: The name of the analyst.
- **Message**: The feedback message.
- **Status**: Either "Approved" (for generated emails) or "Rejected" (for deleted feedbacks).
#### Duplicate Prevention:
- Before processing a feedback (approving or generating an email), the system checks the log to ensure the same ticket ID for the same account has not already been processed.
- If a duplicate is found, the feedback is excluded from the list.

## 🔄 Refreshing Feedback Data
If the Excel files in the **data/** folder are updated while the application is running, you can reload the data without restarting the application.

#### Steps:
1. Update or add Excel files in the **data/** folder.
2. Click the "Refresh Data" button in the main window.
3. The application will reload all feedback data, update the account dropdown, and refresh the displayed feedbacks.
## 🔒 Notes
- This app is designed for manual, local use on a Windows machine with Microsoft Outlook installed.
- Analyst emails must be manually registered in **config/analysts.json**.
- Excel files in the **data/** folder must be named exactly like the account name (e.g., **Flowserve.xlsx** for the "Flowserve" account).
- The application does not send emails automatically; it only creates drafts in Outlook.
- The interface supports pagination to handle large datasets (20 feedbacks per page).
## 📌 Future Ideas
- **Feedback Selection**: Allow users to select specific feedbacks for email generation (currently processes all displayed feedbacks).
- **Confirmation Prompt**: Add a confirmation prompt before generating email drafts.
- **Loading Indicator**: Enhance the loading indicator with a progress bar for long operations.
- **Settings UI Enhancements**: Add more settings options (e.g., manage accounts and groups via the UI).
- **Automatic Feedback Import**: Import feedback data directly from Power BI, an email inbox, or a database.
- **Scheduled Email Dispatch**: Add a feature to schedule weekly email sending.
- **Standalone Executable**: Package the app as a standalone **.exe** using tools like PyInstaller.
- **Error Logging**: Implement detailed error logging to a file for easier debugging.
- **Theming Support**: Add support for light/dark themes configurable via settings.
- **Export Reports**: Allow exporting detailed reports (e.g., all approved feedbacks) to Excel or PDF.
## 🧑‍💻 Developed With
- **Tkinter**: For the graphical user interface (GUI).
- **openpyxl**: For reading and writing Excel files.
- **win32com.client**: For Outlook integration (generating email drafts).
- **Python 3.11+**: Core programming language.
## 🤝 Contributors
- Ribas, Guilherme
## 📃 License
MIT License — feel free to use and adapt!
## Additional Notes for Developers
- The application uses a modular structure with separate files for UI logic (**ui.py**), Excel processing (**processor.py**), and utility functions (**utils.py**).
- The **settings_window.py** file contains the logic for the settings window, allowing configuration of the default sender email and template language.
- The **templates/** folder contains HTML email templates and assets (images) used in the email drafts.
- To extend the application, consider adding new features to the **utils.py** file for reusable functions or enhancing the UI in **ui.py**.
