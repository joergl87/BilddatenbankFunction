import logging
import os
import tempfile
import datetime
import requests

import azure.functions as func
from azure.identity import CertificateCredential, ManagedIdentityCredential
from azure.keyvault.secrets import SecretClient
from azure.mgmt.resource import ResourceManagementClient

# Create the FunctionApp object (Python v2 model)
app = func.FunctionApp()

# --- Helpers ---

def _get_pem_from_kv() -> str:
    """Fetch PEM private key from Key Vault and save to a temp file."""
    try:
        kv_name = os.environ.get("KEYVAULT_NAME")
        secret_name = os.environ.get("SECRET_NAME")

        if not kv_name:
            logging.error("KEYVAULT_NAME environment variable is not set")
            raise ValueError("KEYVAULT_NAME environment variable is not set")
        if not secret_name:
            logging.error("SECRET_NAME environment variable is not set")
            raise ValueError("SECRET_NAME environment variable is not set")

        logging.info(f"Fetching secret '{secret_name}' from Key Vault '{kv_name}'")

        # Auth into Key Vault using the Function App's Managed Identity
        logging.info("Creating ManagedIdentityCredential...")
        credential = ManagedIdentityCredential()

        kv_url = f"https://{kv_name}.vault.azure.net"
        logging.info(f"Creating SecretClient for URL: {kv_url}")
        client = SecretClient(vault_url=kv_url, credential=credential)

        logging.info(f"Attempting to get secret '{secret_name}'...")
        secret = client.get_secret(secret_name)
        logging.info(f"Successfully retrieved secret '{secret_name}', length: {len(secret.value)} chars")

        tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pem")
        tmp_file.write(secret.value.encode("utf-8"))
        tmp_file.close()
        logging.info(f"Saved PEM to temp file: {tmp_file.name}")
        return tmp_file.name
    except Exception as e:
        logging.error(f"Failed to get PEM from Key Vault: {type(e).__name__}: {str(e)}")
        raise


def _get_sp_credential() -> CertificateCredential:
    """Build a credential object for the Service Principal using PEM cert from KV."""
    try:
        tenant_id = os.environ.get("TENANT_ID")
        client_id = os.environ.get("CLIENT_ID")

        if not tenant_id:
            logging.error("TENANT_ID environment variable is not set")
            raise ValueError("TENANT_ID environment variable is not set")
        if not client_id:
            logging.error("CLIENT_ID environment variable is not set")
            raise ValueError("CLIENT_ID environment variable is not set")

        logging.info(f"Creating Service Principal credential - Tenant: {tenant_id}, Client: {client_id}")
        pem_path = _get_pem_from_kv()

        logging.info(f"Creating CertificateCredential with PEM file: {pem_path}")
        credential = CertificateCredential(
            tenant_id=tenant_id,
            client_id=client_id,
            certificate_path=pem_path
        )
        logging.info("CertificateCredential created successfully")
        return credential
    except Exception as e:
        logging.error(f"Failed to create SP credential: {type(e).__name__}: {str(e)}")
        raise


def _login_to_deck() -> None:
    """Login to DECK API using credentials from Key Vault."""
    try:
        kv_name = os.environ.get("KEYVAULT_NAME")

        logging.info("Fetching DECK credentials from Key Vault...")

        # Auth into Key Vault using the Function App's Managed Identity
        credential = ManagedIdentityCredential()
        client = SecretClient(vault_url=f"https://{kv_name}.vault.azure.net", credential=credential)

        # Fetch username and password
        username_secret = client.get_secret("beeboard-username")
        password_secret = client.get_secret("beeboard-password")

        username = username_secret.value
        password = password_secret.value

        logging.info(f"Retrieved credentials for user: {username}")

        # Prepare login request
        login_url = "https://instone.beeboard.eu/gateway/api/v1/login"
        login_data = {
            "grant_type": "password",
            "username": username,
            "password": password
        }

        logging.info(f"Attempting to login to DECK API at: {login_url}")

        # Make login request
        response = requests.post(
            login_url,
            data=login_data,
            headers={"Content-Type": "application/x-www-form-urlencoded"}
        )

        if response.status_code == 200:
            token_data = response.json()
            logging.info(f"Successfully logged in to DECK API. Access token received.")
            if 'expires_in' in token_data:
                logging.info(f"Token expires in: {token_data['expires_in']} seconds")
        else:
            logging.error(f"Failed to login to DECK API. Status code: {response.status_code}")
            logging.error(f"Response: {response.text}")

    except Exception as e:
        logging.error(f"Failed to login to DECK: {type(e).__name__}: {str(e)}")
        raise


def _get_all_items_recursive(drive_id: str, item_id: str, headers: dict, path: str = "") -> list[dict]:
    """Recursively get all items from a folder and its subfolders."""
    all_items = []

    # Get children of current folder
    if item_id == "root":
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root/children?$expand=listItem($expand=fields)"
    else:
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/children?$expand=listItem($expand=fields)"

    logging.info(f"Fetching items from path: {path or 'root'}")

    response = requests.get(url, headers=headers)

    if response.status_code != 200:
        logging.error(f"Failed to get items from {path or 'root'}. Status: {response.status_code}")
        logging.error(f"Response: {response.text}")
        return all_items

    data = response.json()
    items = data.get("value", [])

    for item in items:
        item_name = item.get("name", "")
        current_path = f"{path}/{item_name}" if path else item_name

        # Add path to item for tracking
        item["_fullPath"] = current_path
        all_items.append(item)

        # If it's a folder, recurse into it
        if "folder" in item:
            child_items = _get_all_items_recursive(drive_id, item["id"], headers, current_path)
            all_items.extend(child_items)

    return all_items


def _list_sharepoint_files() -> list[dict]:
    """List all files in the SharePoint document library with their properties."""
    try:
        # Hardcoded SharePoint site and library
        sharepoint_url = "https://instonegmbh.sharepoint.com.mcas.ms/sites/ProjektbilderFuerAnwendungen/beeboardProjektbilder"

        logging.info(f"Fetching files from SharePoint document library: {sharepoint_url}")

        cred = _get_sp_credential()

        # Get access token for Microsoft Graph API
        token = cred.get_token("https://graph.microsoft.com/.default")

        headers = {
            "Authorization": f"Bearer {token.token}",
            "Accept": "application/json"
        }

        # First, we need to resolve the site ID from the URL
        # The actual SharePoint hostname without MCAS proxy
        hostname = "instonegmbh.sharepoint.com"
        site_path = "/sites/ProjektbilderFuerAnwendungen"
        library_name = "beeboard-Projektbilder"

        # Get the site ID
        site_url = f"https://graph.microsoft.com/v1.0/sites/{hostname}:{site_path}"
        logging.info(f"Resolving site ID from: {site_url}")

        site_response = requests.get(site_url, headers=headers)

        if site_response.status_code != 200:
            logging.error(f"Failed to resolve site. Status code: {site_response.status_code}")
            logging.error(f"Response: {site_response.text}")
            return []

        site_data = site_response.json()
        site_id = site_data.get("id")
        logging.info(f"Successfully resolved site ID: {site_id}")

        # Get the document library (drive) by name
        drives_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
        logging.info(f"Fetching drives from: {drives_url}")

        drives_response = requests.get(drives_url, headers=headers)

        if drives_response.status_code != 200:
            logging.error(f"Failed to get drives. Status code: {drives_response.status_code}")
            logging.error(f"Response: {drives_response.text}")
            return []

        drives_data = drives_response.json()
        drives = drives_data.get("value", [])

        # Find the specific library
        drive_id = None
        for drive in drives:
            if drive.get("name") == library_name:
                drive_id = drive.get("id")
                logging.info(f"Found library '{library_name}' with drive ID: {drive_id}")
                break

        if not drive_id:
            logging.error(f"Could not find library '{library_name}' in available drives")
            logging.info(f"Available drives: {[d.get('name') for d in drives]}")
            return []

        # Recursively get all items from the library
        logging.info("Starting recursive fetch of all items...")
        all_items = _get_all_items_recursive(drive_id, "root", headers)

        logging.info(f"Successfully retrieved {len(all_items)} total items from document library")

        # Process and log each file with its properties
        files_info = []
        folder_count = 0
        file_count = 0

        for item in all_items:
            if "folder" in item:
                folder_count += 1
                logging.info(f"Folder: {item.get('_fullPath')}")
            elif "file" in item:  # Only process files, not folders
                file_count += 1
                file_info = {
                    "name": item.get("name"),
                    "path": item.get("_fullPath"),
                    "id": item.get("id"),
                    "size": item.get("size"),
                    "webUrl": item.get("webUrl"),
                    "createdDateTime": item.get("createdDateTime"),
                    "lastModifiedDateTime": item.get("lastModifiedDateTime"),
                }

                # Add custom fields if available
                if "listItem" in item and "fields" in item["listItem"]:
                    file_info["customFields"] = item["listItem"]["fields"]

                files_info.append(file_info)

                # Log detailed information
                logging.info(f"\n--- File #{file_count}: {file_info['name']} ---")
                logging.info(f"  Path: {file_info['path']}")
                logging.info(f"  ID: {file_info['id']}")
                logging.info(f"  Size: {file_info['size']} bytes")
                logging.info(f"  Created: {file_info['createdDateTime']}")
                logging.info(f"  Modified: {file_info['lastModifiedDateTime']}")
                logging.info(f"  URL: {file_info['webUrl']}")

                if "customFields" in file_info:
                    logging.info(f"  Custom Properties:")
                    for key, value in file_info["customFields"].items():
                        logging.info(f"    {key}: {value}")

        logging.info(f"\n=== Summary: Found {file_count} files and {folder_count} folders ===")

        return files_info

    except Exception as e:
        logging.error(f"Failed to list SharePoint files: {type(e).__name__}: {str(e)}")
        logging.exception("Full exception details:")
        return []


def _test_call() -> list[str]:
    """Test Azure auth by listing resource groups or fetching an access token."""
    try:
        subscription_id = os.environ.get("AZURE_SUBSCRIPTION_ID")
        logging.info(f"AZURE_SUBSCRIPTION_ID: {subscription_id if subscription_id else 'NOT SET'}")

        cred = _get_sp_credential()

        if subscription_id:
            logging.info(f"Testing with ResourceManagementClient for subscription: {subscription_id}")
            rm = ResourceManagementClient(credential=cred, subscription_id=subscription_id)
            logging.info("Attempting to list resource groups...")
            rg_list = [rg.name for rg in rm.resource_groups.list()]
            logging.info(f"Successfully listed {len(rg_list)} resource groups: {rg_list}")
            return rg_list
        else:
            logging.info("No subscription ID, testing by getting access token...")
            token = cred.get_token("https://management.azure.com/.default")
            expiry_time = datetime.datetime.fromtimestamp(token.expires_on)
            result = f"Token OK, expires {expiry_time.isoformat()}Z"
            logging.info(f"Token acquisition successful: {result}")
            return [result]
    except Exception as e:
        logging.error(f"Test call failed: {type(e).__name__}: {str(e)}")
        raise

# --- Function itself ---

@app.timer_trigger(
    schedule="0 */5 * * * *",   # every 5 minutes (UTC)
    arg_name="myTimer",
    run_on_startup=True,        # run once on cold start (useful for testing)
    use_monitor=True            # keep track of missed runs
)
def beeboard_image_sync(myTimer: func.TimerRequest) -> None:
    logging.info("===== Azure Function Timer Triggered =====")
    logging.info(f"Timer executed at {datetime.datetime.utcnow().isoformat()}Z")

    # Log all relevant environment variables (without exposing sensitive values)
    env_vars = ["KEYVAULT_NAME", "SECRET_NAME", "TENANT_ID", "CLIENT_ID", "AZURE_SUBSCRIPTION_ID"]
    for var in env_vars:
        value = os.environ.get(var)
        logging.info(f"Environment variable {var}: {'SET' if value else 'NOT SET'}")

    if myTimer.past_due:
        logging.warning("The timer is past due!")

    try:
        logging.info("Starting authentication test...")
        result = _test_call()
        logging.info(f"Auth OK. Test result: {result}")

        # List SharePoint files
        logging.info("Listing files from SharePoint document library...")
        _list_sharepoint_files()

        # If auth test succeeded, try to login to DECK
        logging.info("Starting DECK API login...")
        _login_to_deck()

    except Exception as e:
        logging.error(f"Auth/Test failed with {type(e).__name__}: {str(e)}")
        logging.exception("Full exception details:")

    logging.info("===== Function execution completed =====")
