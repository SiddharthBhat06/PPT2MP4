import os
import shutil
import time
import requests
import streamlit as st
from msal import PublicClientApplication
import win32com.client
import pythoncom
# Environment variables for Azure
CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
TENANT_ID = os.getenv("AZURE_TENANT_ID")
SCOPES = ["https://graph.microsoft.com/Files.Read.All"]
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

if not CLIENT_ID or not TENANT_ID:
    st.error("AZURE_CLIENT_ID and AZURE_TENANT_ID must be set in environment variables.")
    st.stop()

app = PublicClientApplication(CLIENT_ID, authority=AUTHORITY)

# Define folders
base_downloads = os.path.join(os.path.expanduser("~"), "Downloads")
to_run_folder = os.path.join(base_downloads, "Input folder")
processed_folder = os.path.join(to_run_folder, "Archives")
os.makedirs(processed_folder, exist_ok=True)

# Authentication manager
@st.cache_resource
class AuthManager:
    def __init__(self):
        self.flow = None
        self.token = None

    def initiate_device_flow(self):
        self.flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in self.flow:
            raise Exception(f"Device flow initiation failed: {self.flow}")
        return self.flow["message"]

    def acquire_token(self):
        if not self.flow:
            raise Exception("Device flow not initiated.")
        result = app.acquire_token_by_device_flow(self.flow)
        if "access_token" not in result:
            raise Exception("Authentication failed.")
        self.token = result["access_token"]
        return self.token

# List and download files from shared folder
def list_and_download_files(token, foldname, local_dir):
    headers = {"Authorization": f"Bearer {token}"}
    os.makedirs(local_dir, exist_ok=True)
    shared_url = "https://graph.microsoft.com/v1.0/me/drive/sharedWithMe"
    response = requests.get(shared_url, headers=headers)
    response.raise_for_status()
    vsb_folder = None
    for item in response.json().get("value", []):
        if item["name"] == foldname and "folder" in item:
            vsb_folder = item["remoteItem"]
            break
    if not vsb_folder:
        raise Exception(f"{foldname} folder not found in shared items")
    drive_id = vsb_folder["parentReference"]["driveId"]
    item_id = vsb_folder["id"]
    children_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/children"
    response = requests.get(children_url, headers=headers)
    response.raise_for_status()
    files = []
    for item in response.json().get("value", []):
        if "file" in item:
            download_url = item["@microsoft.graph.downloadUrl"]
            local_path = os.path.join(local_dir, item["name"])
            with requests.get(download_url, stream=True) as r:
                with open(local_path, "wb") as f:
                    for chunk in r.iter_content(chunk_size=8192):
                        f.write(chunk)
            files.append(local_path)
    return files

# PowerPoint to MP4 conversion
def ppt_to_mp4(ppt_path, output_mp4_path=None, slide_duration=10, height=720, fps=30, quality=1):
    if not os.path.exists(ppt_path):
        raise FileNotFoundError(f"[❌] File not found: {ppt_path}")
    if not ppt_path.lower().endswith((".pptx", ".pptm")):
        raise ValueError("[❌] File must be a .pptx or .pptm PowerPoint file")
    if output_mp4_path is None:
        output_mp4_path = os.path.splitext(ppt_path)[0] + ".mp4"
    ppt = win32com.client.Dispatch("PowerPoint.Application")
    ppt.Visible = True
    presentation = ppt.Presentations.Open(ppt_path, WithWindow=True)
    presentation.CreateVideo(output_mp4_path, -1, slide_duration, height, fps, quality)
    while True:
        status = presentation.CreateVideoStatus
        if status == 3:
            break
        elif status == 0:
            raise Exception("Export failed.")
        time.sleep(2)
    presentation.Close()
    ppt.Quit()

# Streamlit UI
def main():
    pythoncom.CoInitialize()
    st.title("PowerPoint to MP4 Converter with Microsoft Graph Authentication")

    auth_manager = AuthManager()

    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if "token" not in st.session_state:
        st.session_state.token = None

    if not st.session_state.authenticated:
        st.write("### Step 1: Authenticate")
        if st.button("Start Authentication"):
            try:
                message = auth_manager.initiate_device_flow()
                st.session_state.device_flow_message = message
                st.session_state.flow = auth_manager.flow
            except Exception as e:
                st.error(f"Authentication initiation failed: {e}")

        if "device_flow_message" in st.session_state:
            st.info(st.session_state.device_flow_message)
            if st.button("Complete Authentication"):
                try:
                    token = auth_manager.acquire_token()
                    st.session_state.token = token
                    st.session_state.authenticated = True
                    st.success("Authentication successful!")
                except Exception as e:
                    st.error(f"Authentication failed: {e}")

    if st.session_state.authenticated:
        st.write("### Step 2: Enter Parameters")
        foldname = st.text_input("Shared Folder Name", key="foldname")
        output_folder = st.text_input("Output Folder (absolute path)", key="output_folder")

        if st.button("Run Download and Conversion"):
            if not foldname or not output_folder:
                st.error("Please enter both shared folder name and output folder path.")
            else:
                try:
                    with st.spinner("Downloading files..."):
                        files = list_and_download_files(st.session_state.token, foldname, to_run_folder)
                    st.success(f"Downloaded {len(files)} files.")

                    os.makedirs(output_folder, exist_ok=True)

                    processed_files = []
                    for ppt_file in files:
                        name = os.path.splitext(os.path.basename(ppt_file))[0]
                        output_mp4 = os.path.join(output_folder, f"{name}.mp4")
                        with st.spinner(f"Converting {name} to MP4..."):
                            ppt_to_mp4(ppt_file, output_mp4)
                        processed_files.append(output_mp4)
                        shutil.move(ppt_file, os.path.join(processed_folder, os.path.basename(ppt_file)))
                    st.success(f"Converted {len(processed_files)} files to MP4.")

                    st.write("### Download MP4 files")
                    for mp4_file in processed_files:
                        with open(mp4_file, "rb") as f:
                            st.download_button(label=f"Download {os.path.basename(mp4_file)}", data=f, file_name=os.path.basename(mp4_file))

                except Exception as e:
                    st.error(f"Error: {e}")
    pythoncom.CoUninitialize()

if __name__ == "__main__":
    main()
