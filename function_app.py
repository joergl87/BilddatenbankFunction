import logging
import os
import tempfile
import datetime
import requests

# Configure logging BEFORE importing Azure modules
logging.getLogger('azure').setLevel(logging.WARNING)
logging.getLogger('azure.identity').setLevel(logging.WARNING)
logging.getLogger('azure.core').setLevel(logging.WARNING)
logging.getLogger('msal').setLevel(logging.WARNING)
logging.getLogger('urllib3').setLevel(logging.WARNING)

import azure.functions as func
from azure.identity import CertificateCredential, ManagedIdentityCredential
from azure.keyvault.secrets import SecretClient

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


def _login_to_deck() -> str:
    """Login to DECK API using credentials from Key Vault. Returns access token."""
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
            access_token = token_data.get("access_token")
            logging.info(f"Successfully logged in to DECK API. Access token received.")
            if 'expires_in' in token_data:
                logging.info(f"Token expires in: {token_data['expires_in']} seconds")
            return access_token
        else:
            logging.error(f"Failed to login to DECK API. Status code: {response.status_code}")
            logging.error(f"Response: {response.text}")
            raise Exception(f"Failed to login to DECK API: {response.status_code}")

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

                # Log only essential information
                projekt_nr = file_info.get("customFields", {}).get("Projektnummer0", "N/A")
                url = file_info.get("customFields", {}).get("URL", "N/A")
                logging.info(f"Found: {file_info['name']} | Project: {projekt_nr} | Has URL: {'Yes' if url != 'N/A' else 'No'}")

        logging.info(f"\n=== Summary: Found {file_count} files and {folder_count} folders ===")

        return files_info

    except Exception as e:
        logging.error(f"Failed to list SharePoint files: {type(e).__name__}: {str(e)}")
        logging.exception("Full exception details:")
        return []


def _fetch_beeboard_projects(access_token: str) -> list[dict]:
    """Fetch all projects from beeboard API."""
    try:
        projects_url = "https://instone.beeboard.eu/gateway/api/v1/projects"

        headers = {
            "Authorization": f"Bearer {access_token}",
            "Accept": "application/json"
        }

        logging.info(f"Fetching projects from beeboard API: {projects_url}")

        response = requests.get(projects_url, headers=headers)

        if response.status_code == 200:
            projects = response.json()
            logging.info(f"Successfully fetched {len(projects)} projects from beeboard")
            return projects
        else:
            logging.error(f"Failed to fetch projects. Status code: {response.status_code}")
            logging.error(f"Response: {response.text}")
            return []

    except Exception as e:
        logging.error(f"Failed to fetch beeboard projects: {type(e).__name__}: {str(e)}")
        logging.exception("Full exception details:")
        return []


def _update_project_image(access_token: str, project_id: str, image_url: str) -> bool:
    """Update a project's image URL in beeboard."""
    try:
        update_url = f"https://instone.beeboard.eu/gateway/api/v1/projects/{project_id}/details"

        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }

        payload = {
            "project_image_url": image_url
        }

        logging.info(f"Updating project {project_id} with image URL: {image_url}")

        response = requests.patch(update_url, json=payload, headers=headers)

        if response.status_code in [200, 204]:
            logging.info(f"Successfully updated project {project_id}")
            return True
        else:
            logging.error(f"Failed to update project {project_id}. Status code: {response.status_code}")
            logging.error(f"Response: {response.text}")
            return False

    except Exception as e:
        logging.error(f"Failed to update project {project_id}: {type(e).__name__}: {str(e)}")
        logging.exception("Full exception details:")
        return False


def _sync_images_to_projects(sharepoint_files: list[dict], access_token: str) -> None:
    """Match SharePoint files with beeboard projects and update project images."""
    try:
        # Fetch all projects from beeboard
        projects = _fetch_beeboard_projects(access_token)

        if not projects:
            logging.warning("No projects fetched from beeboard, cannot sync images")
            return

        # Build a mapping of project number -> project
        project_map = {}
        for project in projects:
            project_number = project.get("number")
            if project_number:
                project_map[str(project_number)] = project
            else:
                logging.debug(f"Skipping project without number: {project.get('id', 'Unknown')}")

        logging.info(f"Built project map with {len(project_map)} projects that have numbers")

        # Match SharePoint files with projects
        matched_count = 0
        updated_count = 0
        skipped_count = 0

        for sp_file in sharepoint_files:
            file_name = sp_file.get("name", "Unknown")
            custom_fields = sp_file.get("customFields", {})

            # Get Projektnummer0 and URL from SharePoint
            projekt_nummer = custom_fields.get("Projektnummer0")
            image_url = custom_fields.get("URL")

            # Skip if no project number or URL
            if not projekt_nummer:
                logging.debug(f"Skipping file {file_name}: No Projektnummer0")
                skipped_count += 1
                continue

            if not image_url:
                logging.debug(f"Skipping file {file_name}: No URL field")
                skipped_count += 1
                continue

            # Convert to string for comparison
            projekt_nummer_str = str(projekt_nummer)

            # Find matching project
            if projekt_nummer_str in project_map:
                matched_count += 1
                project = project_map[projekt_nummer_str]
                project_id = project.get("id")
                project_name = project.get("title", "Unknown")
                current_image_url = project.get("project_image_url", "")

                # Check if the URLs are different
                if current_image_url == image_url:
                    logging.info(f"Matched: {projekt_nummer_str} ({project_name}) - already up to date")
                else:
                    logging.info(f"Matched: {projekt_nummer_str} ({project_name}) - updating image URL")
                    # Update the project
                    success = _update_project_image(access_token, project_id, image_url)
                    if success:
                        updated_count += 1
            else:
                skipped_count += 1

        logging.info(f"\n=== Sync Summary ===")
        logging.info(f"Total SharePoint files processed: {len(sharepoint_files)}")
        logging.info(f"Matched with projects: {matched_count}")
        logging.info(f"Successfully updated: {updated_count}")
        logging.info(f"Skipped (no match or missing data): {skipped_count}")

    except Exception as e:
        logging.error(f"Failed to sync images to projects: {type(e).__name__}: {str(e)}")
        logging.exception("Full exception details:")


# --- Function itself ---

@app.timer_trigger(
    schedule="0 0 0,12 * * *",  # twice daily at midnight and noon (UTC)
    arg_name="myTimer",
    run_on_startup=True,        # run once on cold start (useful for testing)
    use_monitor=True            # keep track of missed runs
)
def beeboard_image_sync(myTimer: func.TimerRequest) -> None:
    logging.info("===== Azure Function Timer Triggered =====")
    logging.info(f"Timer executed at {datetime.datetime.utcnow().isoformat()}Z")

    # Log all relevant environment variables (without exposing sensitive values)
    env_vars = ["KEYVAULT_NAME", "SECRET_NAME", "TENANT_ID", "CLIENT_ID"]
    for var in env_vars:
        value = os.environ.get(var)
        logging.info(f"Environment variable {var}: {'SET' if value else 'NOT SET'}")

    if myTimer.past_due:
        logging.warning("The timer is past due!")

    try:
        # Login to DECK API first to get access token
        logging.info("Starting DECK API login...")
        access_token = _login_to_deck()

        # List SharePoint files
        logging.info("Listing files from SharePoint document library...")
        sharepoint_files = _list_sharepoint_files()

        # Sync images to beeboard projects
        if sharepoint_files and access_token:
            logging.info("Starting image sync to beeboard projects...")
            _sync_images_to_projects(sharepoint_files, access_token)
        else:
            logging.warning("Skipping sync: No SharePoint files or no access token")

    except Exception as e:
        logging.error(f"Auth failed with {type(e).__name__}: {str(e)}")
        logging.exception("Full exception details:")

    logging.info("===== Function execution completed =====")
