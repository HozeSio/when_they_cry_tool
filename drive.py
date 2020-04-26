from __future__ import print_function
import pickle
import os.path
import sys
import hashlib
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from apiclient import errors
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/drive']
TRANSLATION_FOLDER_ID = '1Q8BO4CB6tGk-hpYsPOq_Tc6FVPYqP5JA'
#TRANSLATION_FOLDER_ID = '1W7Yxvl3WRzZ1fDbuim8EPediY0qrxWBe'


def get_files(service, folderId, files, filter_folder_name):
    page_token = None
    while True:
        try:
            param = {}
            if page_token:
                param['pageToken'] = page_token
            children = service.files().list(
                fields='files(id, name, mimeType, md5Checksum)',
                q=f"'{folderId}' in parents and trashed = false",
                **param).execute()

            for child in children['files']:
                mimeType = child['mimeType']
                if mimeType == 'application/vnd.google-apps.folder':
                    sub_folder_name = child['name']
                    print(f"searching {sub_folder_name}")
                    if filter_folder_name and sub_folder_name != filter_folder_name:
                        continue
                    files[sub_folder_name] = {}
                    get_files(service, child['id'], files[sub_folder_name], None)
                # xlsx
                elif mimeType == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
                    present_files = files.get('.', [])
                    present_files.append(child)
                    files['.'] = present_files
                elif mimeType == 'text/plain':
                    pass
                else:
                    print(f"unexpected mimeType {mimeType} found, {child['name']}", file=sys.stderr)
            page_token = children.get('nextPageToken')
            if not page_token:
                break
        except errors.HttpError as error:
            print(f'An error occured: {error}')
            break


def get_creds():
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    return creds


def download_folder(drive_service, tree, folder_path):
    downloaders = []
    for folder_name, contents in tree.items():
        parent_folder = os.path.normpath(os.path.join(folder_path, folder_name))
        if folder_name != '.':
            downloaders.extend(download_folder(drive_service, contents, parent_folder))
            continue

        for file in contents:
            if not os.path.exists(parent_folder):
                os.mkdir(parent_folder)
            local_file_path = os.path.join(parent_folder, file['name'])
            if os.path.exists(local_file_path):
                with open(local_file_path, 'rb') as local_file_fd:
                    local_md5 = hashlib.md5(local_file_fd.read()).hexdigest()
                    if local_md5 == file['md5Checksum']:
                        continue

            print(f"Downloading {file['name']} at {local_file_path}")

            request = drive_service.files().get_media(fileId=file['id'])
            fd = open(local_file_path, 'wb')
            downloaders.append((MediaIoBaseDownload(fd, request), fd))
    return downloaders


def upload_folder(drive_service, tree, folder_path):
    for folder_name, contents in tree.items():
        parent_folder = os.path.normpath(os.path.join(folder_path, folder_name))
        if folder_name != '.':
            upload_folder(drive_service, contents, parent_folder)
            continue

        for file in contents:
            local_file_path = os.path.join(parent_folder, file['name'])
            if not os.path.exists(local_file_path):
                print(f"{local_file_path} not exist")
                continue

            with open(local_file_path, 'rb') as local_file_fd:
                local_md5 = hashlib.md5(local_file_fd.read()).hexdigest()
                if local_md5 == file['md5Checksum']:
                    continue

            print(f"Uploading {local_file_path}")
            file = drive_service.files().update(fileId=file['id'],
                                                media_body=MediaFileUpload(local_file_path)
                                                ).execute()


def download_drive(local_folder, filter_folder_name=None):
    creds = get_creds()
    drive_service = build('drive', 'v3', credentials=creds)

    root = {}
    get_files(drive_service, TRANSLATION_FOLDER_ID, root, filter_folder_name)

    downloaders = download_folder(drive_service, root, local_folder)
    while downloaders:
        for item in downloaders[:10]:
            down, fd = item
            try:
                status, done = down.next_chunk()
            except errors.HttpError:
                print(f"Failed to downloading {fd.name}")
                raise
            if done:
                fd.close()
                downloaders.remove(item)


def upload_drive(local_folder, filter_folder_name=None):
    creds = get_creds()
    drive_service = build('drive', 'v3', credentials=creds)

    root = {}
    get_files(drive_service, TRANSLATION_FOLDER_ID, root, filter_folder_name)

    upload_folder(drive_service, root, local_folder)


if __name__ == '__main__':
    if sys.argv[1] == 'download':
        download_drive(f"{os.path.pardir}{os.path.sep}Drive", sys.argv[2] if len(sys.argv) >= 3 else None)
    elif sys.argv[1] == 'upload':
        upload_drive(f"{os.path.pardir}{os.path.sep}Drive", sys.argv[2] if len(sys.argv) >= 3 else None)