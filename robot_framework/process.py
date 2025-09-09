"""This module contains the main process of the robot."""

from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.database.queues import QueueElement

from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext

import requests
import uuid
import pandas as pd

import os


# pylint: disable-next=unused-argument
def process(orchestrator_connection: OrchestratorConnection, queue_element: QueueElement | None = None) -> None:
    """Do the primary process of the robot."""
    orchestrator_connection.log_trace("Running process.")

    NovaToken = orchestrator_connection.get_credential("KMDAccessToken")
    Secret = orchestrator_connection.get_credential("KMDClientSecret")

    NovaTokenAPI = NovaToken.username
    secret = Secret.password
    id = Secret.username


    sharepoint_site = f"{orchestrator_connection.get_constant("AarhusKommuneSharePoint").value}/Teams/tea-teamsite10168"

    certification = orchestrator_connection.get_credential("SharePointCert")
    api = orchestrator_connection.get_credential("SharePointAPI")
    
    tenant = api.username
    client_id = api.password
    thumbprint = certification.username
    cert_path = certification.password
    
    client = sharepoint_client(tenant, client_id, thumbprint, cert_path, sharepoint_site, orchestrator_connection)


    # Authenticate
    auth_payload = {
        "client_secret": secret,
        "grant_type": "client_credentials",
        "client_id": id,
        "scope": "client"
    }

    sharepoint_folder = "Delte dokumenter/Dokumentsøgning"


    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    response = requests.post(NovaTokenAPI, data=auth_payload, headers=headers)

    response.raise_for_status()
    access_token = response.json().get("access_token")

    # Caseworkers to process
    caseworkers = ["AZX0018", "2GBYGSAG Byggeri"]
    api_url = "https://novaapi.kmd.dk/api/Document/GetList?api-version=2.0-Case"

    for caseWorker in caseworkers:
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {access_token}"
        }
        
        payload = {
            "common": {"transactionId": str(uuid.uuid4()).upper()},
            "paging": {
                "startRow": 0,
                "numberOfRows": 1,
                "calculateTotalNumberOfRows": True
            },
            "title": "*hoveddokument*",
            "caseworker": {
                "kspIdentity": {"racfId" if caseWorker == "AZX0018" else "userKey": caseWorker}
            },
            "acceptReceived": False,
            "getOutput": {"title": False}
        }
        
        response = requests.put(api_url, headers=headers, json=payload)
        
        response.raise_for_status()
        
        response_json = response.json()
        
        # Determine values for Excel
        number_of_rows = response_json.get("pagingInformation", {}).get("numberOfRows", 0)
        total_number_of_rows = response_json.get("pagingInformation", {}).get("totalNumberOfRows", 0)
        
        Antal = str(total_number_of_rows if number_of_rows > 0 else 0)
        Indsendelser = "Ændrede indsendelser" if caseWorker == "AZX0018" else "Nye indsendelser"
        orchestrator_connection.log_info(f'{caseWorker}: {Antal}')
        
        # Create DataFrame and save to Excel
        df = pd.DataFrame([{"Indsendelser": Indsendelser, "Antal": Antal}])
        file_name = f"{caseWorker}.xlsx"
        df.to_excel(file_name, index=False)
        upload_file_to_sharepoint(client, sharepoint_folder, file_name, orchestrator_connection)
        os.remove(file_name)


def sharepoint_client(tenant: str, client_id: str, thumbprint: str, cert_path: str, sharepoint_site_url: str, orchestrator_connection: OrchestratorConnection) -> ClientContext:
    """
    Creates and returns a SharePoint client context.
    """
    # Authenticate to SharePoint
    cert_credentials = {
        "tenant": tenant,
        "client_id": client_id,
        "thumbprint": thumbprint,
        "cert_path": cert_path
    }
    ctx = ClientContext(sharepoint_site_url).with_client_certificate(**cert_credentials)

    # Load and verify connection
    web = ctx.web
    ctx.load(web)
    ctx.execute_query()

    orchestrator_connection.log_info(f"Authenticated successfully. Site Title: {web.properties['Title']}")
    return ctx

def upload_file_to_sharepoint(client: ClientContext, sharepoint_file_url: str, local_file_path: str, orchestrator_connection: OrchestratorConnection):
    """
    Uploads the specified local file back to SharePoint at the given URL.
    Uses the folder path directly to upload files.
    """
    # Extract the root folder, folder path, and file name
    path_parts = sharepoint_file_url.split('/')
    DOCUMENT_LIBRARY = path_parts[0]  # Root folder name (document library)
    FOLDER_PATH = path_parts[1]
    file_name = os.path.basename(local_file_path)  # File name

    # Construct the server-relative folder path (starting with the document library)
    if FOLDER_PATH:
        folder_path = f"{DOCUMENT_LIBRARY}/{FOLDER_PATH}"
    else:
        folder_path = f"{DOCUMENT_LIBRARY}"

    # Get the folder where the file should be uploaded
    target_folder = client.web.get_folder_by_server_relative_url(folder_path)
    client.load(target_folder)
    client.execute_query()

    # Upload the file to the correct folder in SharePoint
    with open(local_file_path, "rb") as file_content:
        uploaded_file = target_folder.upload_file(file_name, file_content).execute_query()