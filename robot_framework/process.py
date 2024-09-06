"""This module uploads data from Excel file to the orchestrator queue."""
import os
import glob
import json
import ast
import uuid
from datetime import datetime
import pandas as pd
import sqlalchemy
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
from mbu_dev_shared_components.utils.fernet_encryptor import Encryptor
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection

# Configuration
SHAREPOINT_SITE_URL = "https://aarhuskommune.sharepoint.com/teams/MBU-RPA-Egenbefordring"
DOCUMENT_LIBRARY = "Delte dokumenter/General/Til udbetaling"


def process(orchestrator_connection: OrchestratorConnection) -> None:
    """Primary process of the robot."""
    orchestrator_connection.log_trace("Running process.")
    process_args = json.loads(orchestrator_connection.process_arguments)
    path_arg = process_args['path']
    naeste_agent_arg = process_args['naeste_agent']

    service_konto_credential = orchestrator_connection.get_credentials("SvcRpaMBU002")
    username = service_konto_credential.username
    password = service_konto_credential.password

    clear_queue(orchestrator_connection)
    fetch_files(username, password, path_arg)
    data_df = load_excel_data(path_arg)
    processed_df = process_data(data_df, naeste_agent_arg)
    approved_df = processed_df[processed_df['is_godkendt']]
    upload_to_queue(approved_df, orchestrator_connection)


def clear_queue(orchestrator_connection: OrchestratorConnection) -> None:
    """Clear elements from the queue."""
    queue_elements = orchestrator_connection.get_queue_elements("Koerselsgodtgoerelse_egenbefordring")
    for element in queue_elements:
        orchestrator_connection.delete_queue_element(element.id)


def fetch_files(username, password, download_path: str) -> None:
    """Download Excel files from SharePoint to the specified path."""
    if not os.path.exists(download_path):
        os.makedirs(download_path)

    ctx = ClientContext(SHAREPOINT_SITE_URL).with_credentials(UserCredential(username, password))
    target_folder_url = f"/teams/MBU-RPA-Egenbefordring/{DOCUMENT_LIBRARY}"
    target_folder = ctx.web.get_folder_by_server_relative_url(target_folder_url)
    files = target_folder.files
    ctx.load(files)
    ctx.execute_query()

    if not files:
        print("No files found in the specified SharePoint folder.")
        return

    for file in files:
        if file.name.endswith('.xlsx'):
            download_path_file = os.path.join(download_path, file.name)
            with open(download_path_file, "wb") as local_file:
                file_content = File.open_binary(ctx, file.serverRelativeUrl)
                local_file.write(file_content.content)
            print(f"Downloaded: {file.name} to {download_path_file}")


def load_excel_data(dir_path: str) -> pd.DataFrame:
    """Load the first Excel file matching the pattern from the directory."""
    print("Loading Excel data...")
    excel_files = glob.glob(os.path.join(dir_path, "Egenbe[forfart]*.xlsx"))
    if not excel_files:
        raise FileNotFoundError("No Egenbefordring... .xlsx files found in the specified folder.")

    file_to_read = excel_files[0]
    df = pd.read_excel(file_to_read)
    print(f"Data loaded from: {file_to_read}")
    os.remove(file_to_read)
    print(f"Deleted: {file_to_read}")
    return df


def extract_url_from_attachments(attachments_str: str) -> str:
    """Extract the URL from the attachments string."""
    if isinstance(attachments_str, str):
        start_index = attachments_str.find('https://')
        if start_index != -1:
            end_index = attachments_str.find("'", start_index)
            if end_index != -1 and end_index > start_index:
                return attachments_str[start_index:end_index]
    return pd.NA


def extract_months_and_year(test_str):
    """Extract months and year from the test string."""
    month_map = {
        'January': 'Januar',
        'February': 'Februar',
        'March': 'Marts',
        'April': 'April',
        'May': 'Maj',
        'June': 'Juni',
        'July': 'Juli',
        'August': 'August',
        'September': 'September',
        'October': 'Oktober',
        'November': 'November',
        'December': 'December'
    }
    data = ast.literal_eval(test_str)
    months = set()
    year = None

    for entry in data:
        if isinstance(entry, dict) and 'dato' in entry:
            date_str = entry['dato']
            date_obj = datetime.strptime(date_str, '%Y-%m-%d')
            month_name = date_obj.strftime('%B')
            months.add(month_map.get(month_name, month_name))
            year = date_obj.year

    sorted_months = sorted(months, key=lambda x: list(month_map.values()).index(x))

    month_str = '/'.join(sorted_months)
    result = f"{month_str} {year}"

    return result


def process_data(df: pd.DataFrame, naeste_agent: str) -> pd.DataFrame:
    """Process the data and return a DataFrame with the required format."""
    encryptor = Encryptor()
    processed_data = []

    for _, row in df.iterrows():
        month_year = extract_months_and_year(row['test'])
        cpr_nr = str(row['cpr_nr_paaanden']) if not pd.isnull(row['cpr_nr_paaanden']) else str(row['cpr_nr'])
        attachments_str = str(row.get('attachments', ''))
        url = extract_url_from_attachments(attachments_str)
        skoleliste = str(row['skoleliste']).lower() if not pd.isnull(row['skoleliste']) else ''

        psp_value = determine_psp_value(skoleliste, row)

        encrypted_cpr = encryptor.encrypt(cpr_nr).decode('utf-8')

        new_row = {
            'cpr_encrypted': encrypted_cpr,
            'beloeb': row['aendret_beloeb_i_alt'] if not pd.isnull(row['aendret_beloeb_i_alt']) else row['beloeb_i_alt'],
            'reference': month_year,
            'arts_konto': '40430002',
            'psp': psp_value,
            'posteringstekst': f"Egenbefordring {month_year}",
            'naeste_agent': naeste_agent,
            'attachment': url,
            'uuid': row.get('uuid', pd.NA),
            'godkendt_af': row.get('godkendt_af', pd.NA),
            'skole': row['skriv_dit_barns_skole_eller_dagtilbud'] if not pd.isnull(row['skriv_dit_barns_skole_eller_dagtilbud']) else row['skoleliste'],
            'is_godkendt': 'x' in str(row.get('godkendt', '')).lower(),
        }

        processed_data.append(new_row)

    return pd.DataFrame(processed_data)


def determine_psp_value(skoleliste: str, row: pd.Series) -> str:
    """Determine PSP value based on school list."""
    if 'langagerskolen' in skoleliste:
        return "XG-5240220808-00004"
    if 'stensagerskolen' in skoleliste:
        return "XG-5240220808-00005"
    if not pd.isnull(row['skriv_dit_barns_skole_eller_dagtilbud']):
        return "XG-5240220835-00004"
    return "XG-5240220808-00003"


def make_unique_references(references: list) -> list:
    """Generate unique references by appending UUIDs."""
    return [f"{ref}_{uuid.uuid4().hex}" for ref in references]


def upload_to_queue(result_df: pd.DataFrame, orchestrator_connection: OrchestratorConnection) -> None:
    """Upload the processed data to the orchestrator queue."""
    queue_data = [json.dumps(data, ensure_ascii=False) for data in result_df.to_dict(orient='records')]
    queue_references = [str(row['posteringstekst']) for _, row in result_df.iterrows()]
    unique_references = make_unique_references(queue_references)

    try:
        print("Uploading data to queue...")
        orchestrator_connection.bulk_create_queue_elements(
            "Koerselsgodtgoerelse_egenbefordring",
            references=unique_references,
            data=queue_data
        )
        print("Data uploaded to queue successfully.")

    except sqlalchemy.exc.IntegrityError as ie:
        print(f"IntegrityError: {ie.orig}")

    except (ValueError, TypeError) as e:
        print(f"Error occurred: {e}")
