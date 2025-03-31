
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.database.queues import QueueElement
import os
from datetime import datetime, timedelta
import locale
import pandas as pd
import pyodbc
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
import smtplib
from email.message import EmailMessage
from openpyxl.utils import get_column_letter
from openpyxl.styles import numbers
from openpyxl.styles import NamedStyle

def process(orchestrator_connection: OrchestratorConnection):
    def sharepoint_client(username: str, password: str, sharepoint_site_url: str) -> ClientContext:
        """
        Creates and returns a SharePoint client context.
        """
        # Authenticate to SharePoint
        ctx = ClientContext(sharepoint_site_url).with_credentials(UserCredential(username, password))

        # Load and verify connection
        web = ctx.web
        ctx.load(web)
        ctx.execute_query()

        orchestrator_connection.log_info(f"Authenticated successfully")
        return ctx

    def upload_to_sharepoint(client: ClientContext, folder_name: str, file_path: str, folder_url: str):
            """
            Uploads a file to a specific folder in a SharePoint document library.

            :param client: Authenticated SharePoint client context
            :param folder_name: Name of the target folder within the document library
            :param file_path: Local file path to upload
            :param folder_url: SharePoint folder URL where the file should be uploaded
            """
            try:
                # Extract file name safely
                file_name = os.path.basename(file_path)

                # Define the SharePoint document library structure
                document_library = f"{folder_url.split('/', 1)[-1]}/Delte Dokumenter/Aktindsigter"
                folder_path = f"{document_library}/{folder_name}"

                # Read file into memory (Prevents closed file issue)
                with open(file_path, "rb") as file:
                    file_content = file.read()  

                # Get SharePoint folder reference
                target_folder = client.web.get_folder_by_server_relative_url(folder_url)

                # Upload file using byte content
                target_folder.upload_file(file_name, file_content)
                
                # Execute request
                client.execute_query()
                orchestrator_connection.log_info(f"✅ Successfully uploaded: {file_name} to {folder_path}")

            except Exception as e:
                orchestrator_connection.log_info(f"❌ Error uploading file: {str(e)}")

    orchestrator_connection.log_info('Starting process Moesgaardlisten')
    RobotCredentials = orchestrator_connection.get_credential('RobotCredentials')
    OldTimeStamp = orchestrator_connection.get_constant('MoesgaardlistenTimestamp').value
    SharepointUrl = orchestrator_connection.get_constant('AarhusKommuneSharepoint').value

    orchestrator_connection.log_info(f'Time of latest run; {OldTimeStamp}')

    # Setting the datestuff
    locale.setlocale(locale.LC_TIME, "da_DK.UTF-8")
    current_date = datetime.now()
    current_week_number = current_date.isocalendar()[1]
    previous_week_start_date = current_date - timedelta(days=7)
    previous_week_number = previous_week_start_date.isocalendar()[1]

    # orchestrator_connection.log_infoing it
    orchestrator_connection.log_info(f"Current Date: {current_date.strftime("%d/%m/%Y")}")
    orchestrator_connection.log_info(f"Current Week Number: {current_week_number}")
    orchestrator_connection.log_info(f"Previous Week Start Date: {previous_week_start_date.strftime("%d/%m/%Y")}")
    orchestrator_connection.log_info(f"Previous Week Number: {previous_week_number}")

    #Setting date according to weekday
    today = datetime.today()
    current_day = today.weekday()
    calculated_date = today- timedelta(days = current_day + 1)
    DatoTil = calculated_date.strftime("%Y-%m-%d")

    orchestrator_connection.log_info(f'Calculated date: {DatoTil}')

    DatoFra = datetime.fromisoformat(OldTimeStamp).strftime("%Y-%m-%d")
    DagsDato = datetime.now().strftime("%Y%m%d")
    DagsDatoÅr = datetime.now().year
    ExcelFileName = f'{DagsDato} Byggesager fra uge {previous_week_number} {DagsDatoÅr}.xlsx'
    user_name = os.getlogin()
    excel_file_path = f"C:\\Users\\{user_name}\\Downloads\\{ExcelFileName}"
    orchestrator_connection.log_info(excel_file_path)

    # Read the SQL query from file
    sql_file_path = "Moesgaardlisten_Ny.sql"
    with open(sql_file_path, "r", encoding="utf-8") as file:
        query = file.read()

    # Replace placeholders with actual date values, ensuring they are formatted correctly
    query = query.replace("@datoFra", f"{DatoFra}").replace("@datoTil", f"{DatoTil}")  # Add single quotes
    orchestrator_connection.log_info(f"Executing SQL query:\n {query}")

    # Database connection setup
    sql_server = orchestrator_connection.get_constant("SqlServer").value
    conn_str = 'DRIVER={ODBC Driver 17 for SQL Server};' + f'SERVER={sql_server};DATABASE=LOIS;Trusted_Connection=yes'
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()

    # Step 2: Execute the actual SELECT query
    cursor.execute(query)

    # Step 3: Fetch results
    rows = cursor.fetchall()
    orchestrator_connection.log_info(f'rows {rows}')

    # Step 4: Load the data into a Pandas DataFrame
    data = pd.read_sql(query, conn)

    # Step 5: Close database connection
    cursor.close()
    conn.close() 

    # Sikrer at 'Sagsdato' er datetime-type
    data['Sagsdato'] = pd.to_datetime(data['Sagsdato'], errors='coerce')

    antal_sager = len(data)

    if antal_sager > 0:
        max_sagsdato = data['Sagsdato'].max()
        orchestrator_connection.log_info(f'Opretter excelfil {ExcelFileName}')

        with pd.ExcelWriter(excel_file_path, engine='xlsxwriter', datetime_format='dd-mm-yyyy') as writer:
            data.to_excel(writer, sheet_name='Byggesager', index=False)

            workbook  = writer.book
            worksheet = writer.sheets['Byggesager']

            # Lav Excel-tabel over hele området
            (max_row, max_col) = data.shape
            col_range = chr(65 + max_col - 1)
            table_range = f"A1:{col_range}{max_row + 1}"

            worksheet.add_table(table_range, {
                'name': 'ByggesagerTable',
                'columns': [{'header': col} for col in data.columns]
            })

            # Tilpas kolonnebredde
            for i, column in enumerate(data.columns):
                column_length = max(data[column].astype(str).map(len).max(), len(column))
                worksheet.set_column(i, i, column_length + 2)

        orchestrator_connection.log_info('Overfører excelfil til sharepoint')
        # file_url = f'{SharepointUrl}/Teams/sec-lukket1752/Delte%20dokumenter'
        file_url = f'{SharepointUrl}/Teams/tea-teamsite11819/Delte Dokumenter/Testmappe'
        client = sharepoint_client(username=RobotCredentials.username, password=RobotCredentials.password, sharepoint_site_url= f'{SharepointUrl}/Teams/tea-teamsite11819/' )
        upload_to_sharepoint(client= client, folder_name = 'Testmappe', file_path=excel_file_path, folder_url= '/Teams/tea-teamsite11819/Delte Dokumenter/Testmappe')
        orchestrator_connection.log_info(f'Uploaded to {file_url}')

        # SMTP Configuration (from your provided details)
        SMTP_SERVER = "smtp.adm.aarhuskommune.dk"
        SMTP_PORT = 25
        SCREENSHOT_SENDER = "moesgaardlisten@aarhus.dk"
        subject = f"Nye byggesager - Moesgaardlisten"

        # Email body (HTML)
        body = f"""
        <html>
            <body>
                <p>Kære XXX,</p>
                <br>
                <p>Robotten har kørt korrekt, der er: {antal_sager} nye sager i uge: {current_week_number}</p>
                <br>
                <p>Med venlig hilsen</p>
                <br>
                <p>Teknik og Miljø</p>
                <p>Digitalisering</p>
                <p>Aarhus Kommune</p>
            </body>
        </html>
        """
        # Create the email message
        to_address = orchestrator_connection.get_constant('balas').value
        UdviklerMail = to_address
        msg = EmailMessage()
        msg['To'] = ', '.join(to_address) if isinstance(to_address, list) else to_address
        msg['From'] = SCREENSHOT_SENDER
        msg['Subject'] = subject
        msg.set_content("Please enable HTML to view this message.")
        msg.add_alternative(body, subtype='html')
        msg['Reply-To'] = UdviklerMail
        msg['Bcc'] = UdviklerMail

        # Send the email using SMTP
        try:
            with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as smtp:
                smtp.send_message(msg)
        except Exception as e:
            orchestrator_connection.log_info(f"Failed to send success email: {e}")

        orchestrator_connection.update_constant('MoesgaardlistenTimestamp', str(max_sagsdato.strftime('%Y-%m-%d')))

        if os.path.exists(excel_file_path):
            os.remove(excel_file_path)