import os
import requests
from urllib.parse import urlparse
from azure.core.credentials import AzureKeyCredential
from azure.search.documents import SearchClient

def list_indexed_presentations(endpoint, key, index_name):
    """
    Queries the Azure Search Index to get a list of all unique files.
    """
    print("AGENT TOOL (STABLE): Querying Azure Search Index for all known files...")
    try:
        # Initialize an Azure Cognitive Search client
        search_client = SearchClient(endpoint=endpoint, index_name=index_name, credential=AzureKeyCredential(key))
        
        # Query all documents, selecting only 'file_name' and asking for total count
        results = search_client.search(search_text="*", select=["file_name"], include_total_count=True)
        
        processed_files = set()  # Use a set to collect unique file names
        for result in results:
            file_name = result.get("file_name")
            if file_name:
                processed_files.add(file_name)
        
        print(f"AGENT TOOL (STABLE): Found {len(processed_files)} unique files in the index.")
        return list(processed_files)
    except Exception as e:
        # Log any errors and return an empty list on failure
        print(f"AGENT TOOL (STABLE) ERROR: Failed to query search index: {e}")
        return []


def _get_all_files_recursively(token, drive_id, folder_id):
    """
    Performs a deep, recursive search to find all files starting from a root folder.
    Returns a dictionary mapping filenames to their download URLs.
    """
    file_map = {}                      # filename -> download URL
    folders_to_scan = [folder_id]      # queue for BFS over folders
    headers = {"Authorization": f"Bearer {token}"}  # Authorization header for Graph API
    
    while folders_to_scan:
        current_folder_id = folders_to_scan.pop(0)  # Pop next folder to scan
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{current_folder_id}/children"
        try:
            # List children (files/folders) under the current folder
            r = requests.get(url, headers=headers)
            r.raise_for_status()
            items = r.json().get("value", [])
            for item in items:
                if "folder" in item:
                    # Enqueue subfolder for further scanning
                    folders_to_scan.append(item['id'])
                elif "file" in item:
                    # Capture direct download URL when available
                    if '@microsoft.graph.downloadUrl' in item:
                        file_map[item['name']] = item['@microsoft.graph.downloadUrl']
        except Exception as e:
            # Non-fatal: continue scanning remaining folders
            print(f"  - Warning: Could not scan folder {current_folder_id}. Error: {e}")
            
    return file_map


def download_files_from_sharepoint(token, drive_id, folder_id, files_to_download, save_dir):
    """
    Downloads files if they don't already exist in the local cache (save_dir).
    """
    print(f"AGENT TOOL (STABLE): Ensuring {len(files_to_download)} files are cached locally...")
    os.makedirs(save_dir, exist_ok=True)  # Ensure cache directory exists
    
    # Build a map of all files (name -> download URL) from the SharePoint root folder
    sp_file_map = _get_all_files_recursively(token, drive_id, folder_id)

    for file_name in files_to_download:
        local_path = os.path.abspath(os.path.join(save_dir, file_name))
        
        # Skip download if already cached locally
        if os.path.exists(local_path):
            # print(f"  - CACHE HIT for '{file_name}'.")  # Optional: enable for verbose logs
            continue

        print(f"  - CACHE MISS for '{file_name}'. Downloading from SharePoint...")
        if file_name in sp_file_map and sp_file_map[file_name]:
            try:
                download_url = sp_file_map[file_name]
                # Fetch the file content from the pre-authenticated URL
                file_resp = requests.get(download_url)
                file_resp.raise_for_status()
                
                # Write to local cache path
                with open(local_path, "wb") as f:
                    f.write(file_resp.content)
                print(f"  ✅ Download complete. Saved to cache.")

            except Exception as e:
                # Log a failure for this file and continue with others
                print(f"  ❌ Download failed for {file_name}. Error: {e}")
        else:
            # If file name not found in the SharePoint map, skip gracefully
            print(f"  - ⚠️ Could not find '{file_name}' on SharePoint. Skipping.")
            
    print("AGENT TOOL (STABLE): Pre-caching process complete.")


def get_access_token(tenant_id, client_id, client_secret):
    # OAuth 2.0 Client Credentials flow against Azure AD to obtain a Graph token
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default"
    }
    r = requests.post(url, data=data, timeout=30)
    r.raise_for_status()
    return r.json()["access_token"]  # Return the access token string


def get_site_and_drive_id(token, sharepoint_url):
    # Parse the human-friendly SharePoint site URL into hostname and path
    parsed = urlparse(sharepoint_url)
    hostname = parsed.hostname
    site_path = parsed.path.strip("/")
    
    # Resolve the site object from hostname + site path
    site_url = f"https://graph.microsoft.com/v1.0/sites/{hostname}:/{site_path}"
    r_site = requests.get(site_url, headers={"Authorization": f"Bearer {token}"})
    r_site.raise_for_status()
    site_id = r_site.json()["id"]
    
    # Enumerate drives (document libraries) under this site and take the first one
    drive_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
    r_drive = requests.get(drive_url, headers={"Authorization": f"Bearer {token}"})
    r_drive.raise_for_status()
    drive_id = r_drive.json()["value"][0]["id"]
    
    return site_id, drive_id  # Return both IDs for subsequent Graph calls