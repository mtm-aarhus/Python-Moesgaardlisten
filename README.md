Moesgaardlisten Robot
Moesgaardlisten is an automation for Teknik og Miljø, Aarhus Kommune. It queries the building case database, generates a weekly report of new cases, saves an Excel file, uploads it to SharePoint, and emails stakeholders automatically.

🚀 Features
✅ Automated Weekly Reporting

Calculates date intervals dynamically each run

Fetches all new building cases registered since the last run

📄 Excel Report Generation

Exports results to a formatted Excel file

Applies dynamic table formatting and column widths

📤 SharePoint Integration

Uploads the Excel file to a SharePoint document library

Verifies upload success

📧 Email Notifications

Sends an HTML email summarizing the number of new cases

Includes the reporting week in the message

🔐 Credential Management

Securely manages SharePoint and database credentials in OpenOrchestrator

🗑️ Cleanup

Deletes local Excel files after upload

🧭 Process Flow
Date Calculation

Retrieves last run timestamp (MoesgaardlistenTimestamp)

Calculates date range for the current week

SQL Query Execution

Loads query from Moesgaardlisten_Ny.sql

Replaces placeholders @datoFra and @datoTil

Executes query against the LOIS SQL Server

Excel File Creation

Saves all retrieved records to Excel

Formats as a table with column autofit

Upload to SharePoint

Authenticates to SharePoint with robot credentials

Uploads Excel to Delte Dokumenter folder

Notification Email

Sends an email to developers summarizing run results

BCCs a monitoring mailbox

Timestamp Update

Updates MoesgaardlistenTimestamp to the latest case date

File Cleanup

Deletes the Excel file from local storage

🔐 Privacy & Security
All communication uses HTTPS or trusted connections

Credentials are stored securely in OpenOrchestrator

No personal data is stored locally after upload

⚙️ Dependencies
Python 3.10+

pandas

pyodbc

openpyxl

xlsxwriter

office365-rest-python-client

smtplib

👷 Maintainer
Gustav Chatterton
Digital udvikling, Teknik og Miljø, Aarhus Kommune

