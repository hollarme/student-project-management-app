
import streamlit as st
import pandas as pd

import gspread 
from gspread.exceptions import WorksheetNotFound, SpreadsheetNotFound
from gspread_dataframe import get_as_dataframe, set_with_dataframe

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from googleapiclient.http import MediaFileUpload, MediaIoBaseUpload

from getfilelistpy import getfilelist


folder_id = None
parent_folder_id = None

SCOPES=['https://www.googleapis.com/auth/drive', 'https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive.file']

credentials = service_account.Credentials.from_service_account_info(st.secrets["ServiceAccountCredentialsSheet"], scopes = SCOPES)

service = build('drive', 'v3', credentials=credentials)

gc = gspread.authorize(credentials)


# class GoogleDriveService:
#     def __init__(self):
#         self._SCOPES=['https://www.googleapis.com/auth/drive', 'https://spreadsheets.google.com/feeds', 
#                       'https://www.googleapis.com/auth/drive.file']
        
#         # self.ServiceAccountCredentials = st.secrets['ServiceAccountCredentials']
#         # st.write(st.secrets['ServiceAccountCredentials'])
#         # self.jsonString = json.dumps({key:value for key, value in self.ServiceAccountCredentials.items()})
        
#     def build(self):
#         creds = service_account.Credentials.from_service_account_info(st.secrets["ServiceAccountCredentialsSheet"], scopes = self._SCOPES)
#         service = build('drive', 'v3', credentials=creds)

#         return service
    
def getFileListFromGDrive():
    selected_fields="files(id,name,webViewLink)"
    list_file=service.files().list(fields=selected_fields).execute()
    return {"files":list_file.get("files")}


def create_folder(folder_name, parent_folder_id):
    """ Create a folder in the parent folder 
    Returns : new folder Id
    """

    files_to_download = getFileListFromGDrive()

    if folder_name in [file['name'] for file in files_to_download['files']]:
        raise FileExistsError
    
    try:
        # create drive api client
        # service = GoogleDriveService().build()
        file_metadata = {
            'name': f'{folder_name}',
            'parents' : [f'{parent_folder_id}'],
            'mimeType': 'application/vnd.google-apps.folder'
        }

        file = service.files().create(body=file_metadata, fields='id').execute()
        return file.get('id')
    except HttpError as error:
        print(F'An error occurred: {error}')
        return None
    
#     new_permission = {
#           'emailAddress': 'eakinboboye@oauife.edu.ng',
#           'type': 'user',
#           'role': 'writer'
#         }
        
#     try:
#         service.permissions().create(fileId=file.get('id'), body=new_permission, supportsAllDrives=True).execute()
#         return file.get('id')
#     except HttpError as error:
#         print ('An error occurred: %s' % error)
#         return None
    
def upload_to_folder(folder_id, file_name, file):
    """Upload a file to the specified folder and prints file ID, folder ID
    Args: Id of the folder
    Returns: ID of the file uploaded
    """
    
    # create drive api client
    # service = GoogleDriveService().build()
        
    files_uploaded = getFileListFromGDrive()

    if file_name in [file['name'] for file in files_uploaded['files'] if file['name'].endswith('.pdf')]:
        
        file_id = [file['id'] for file in files_uploaded['files'] if file['name']==file_name][0]
        
        service.files().delete(fileId=file_id).execute()
        
        # raise FileExistsError
        
    try:
        
        file_metadata = {
            'name': f'{file_name}',
            'parents': [folder_id]
        }
        media = MediaIoBaseUpload(file,
                                mimetype='application/pdf', resumable=True)
        # pylint: disable=maybe-no-member
        file = service.files().create(body=file_metadata, media_body=media,
                                      fields='id').execute()
        print(F'File ID: "{file.get("id")}".')
        return file.get('id')

    except HttpError as error:
        print(F'An error occurred: {error}')
        return None


def get_database(parent_folder_id, db_name, sheet):
#         # Create a list of scope values to pass to the credentials object
#         scope = ['https://spreadsheets.google.com/feeds',
#                 'https://www.googleapis.com/auth/drive']

#         # Create a credentials object using the service account info and scope values
#         credentials = service_account.Credentials.from_service_account_info(
#                     st.secrets["ServiceAccountCredentialsSheet"], scopes = scope)

        # Authorize the connection to Google Sheets using the credentials object
        # gc = gspread.authorize(credentials)
        
        try:
            # Open the Google Sheets document with the specified name
            sh = gc.open(db_name, parent_folder_id)
        except SpreadsheetNotFound:
            # Create the Google Sheets document with the specified name
            sh = gc.create(db_name, parent_folder_id)
        try:
            # Access the worksheet within the document with the specified name
            worksheet = sh.worksheet(sheet) 
        except WorksheetNotFound:
            #Create the worksheet for the user
            worksheet = sh.add_worksheet(sheet, rows=1000, cols=50)
            
        return worksheet


setup, tab0, tab1, tab2 = st.tabs(["Setup", "Information", "Grouping", "File Upload"])

with setup:
    
    st.header('Setup Page')
    
    session = st.selectbox(
    'What is the current session?',
    ('2021/2022', '2022/2023', '2023/2024', '2024/2025', '2025/2026', '2026/2027', '2027/2028', '2028/2029', '2029/2030'))
    #create folders (with the above session names) manually in eakinboboye@oauife.edu.ng and share with sheet-admin@examtools.iam.gserviceaccount.com
    
    semester = st.selectbox(
    'What is the current semester?',
    ('Rain', 'Harmattan'))
    
    course_name = st.selectbox(
    'What course is it?',
    ('EEE501', 'EEE502'))
    
    parent_folder_name = f"{session}-{semester}-{course_name}"
            
    result = getFileListFromGDrive()
    
    try:
        st.session_state.folder_id = [dic['id'] for dic in result.get('files') if dic['name']==parent_folder_name][0]
    except IndexError:
        if folder_name in ["2021/2022-Rain-EEE501", "2021/2022-Rain-EEE502"]:
            st.session_state.folder_id = '1EJQyD0NghC1lxalJCWINQYvqAZdnGDve'
            

parent_folder_id = st.session_state.folder_id

with tab0:
  
    st.markdown('''
                _`EEE 501 & 502` project management application_
                
                ''')

        
    with st.expander('UPLOADING FILES', expanded=False):
        st.markdown('''
                    You are required to submit the final (corrected) copy of you thesis: use the file upload tab!
                    
                    __Note__: only pdf files are accepted!
                    
                    
                    ''')

with tab1:
    try:
        grouping_tables = get_as_dataframe(get_database(parent_folder_id, "Defense Grouping List", 'Sheet1'),usecols=['Reg. Number','Names','Adviser','Group','Staff']).dropna(how='all')#pd.read_csv('defense_grouping_list.csv', nrows=1000)
        grouping_tables.set_index('Reg. Number',inplace=True)
        grouping_table = [grouping_tables.loc[indices,:] for indices in grouping_tables.groupby('Group').groups.values()]
    except FileNotFoundError:
        grouping_tables = pd.DataFrame({})

    st.write('## Defense Grouping List')
    st.dataframe(grouping_tables)

with tab2:
    file = None
    
    st.write('### Upload a PDF file against your registration number')    

    folder_name = st.selectbox(
    f'Select your registration number:',
    [student for student in grouping_tables.index], key=1)

    # st.write(f'You are in Group {grouping_tables.loc[folder_name].Group if folder_name else 0: 1}')
    try:
        folder_id = create_folder(folder_name, parent_folder_id)
    except FileExistsError:
        all_files = getFileListFromGDrive()
        folder_id = [file['id'] for file in all_files['files'] if file['name']==folder_name][0]
            
    find_topic_list = gc.list_spreadsheet_files(folder_id=parent_folder_id)
    topic_list_exist = any([True if dic['name']=='Topic List' else False for dic in find_topic_list])
    
    # print(topic_list_exist)
    
    if not topic_list_exist:
        get_database(parent_folder_id, 'Topic List', 'Sheet1') #creates the sheet
        topic_dataframe = get_as_dataframe(get_database(parent_folder_id, "Defense Grouping List", 'Sheet1'),usecols=['Reg. Number','Names','Adviser']).dropna(how='all')#pd.read_csv('defense_grouping_list.csv', nrows=1000)
        topic_dataframe.set_index('Reg. Number',inplace=True)
        topic_dataframe['Title'] = ''
        set_with_dataframe(get_database(parent_folder_id, 'Topic List', 'Sheet1'), topic_dataframe, include_index=True)
    
    topic_dataframe = get_as_dataframe(get_database(parent_folder_id, "Topic List", 'Sheet1'),usecols=['Reg. Number','Names','Adviser', 'Title']).dropna(how='all')#pd.read_csv('defense_grouping_list.csv', nrows=1000)
    topic_dataframe.set_index('Reg. Number',inplace=True)
    topic = topic_dataframe.at[folder_name, 'Title']
    
    project_title = st.text_area('Enter the title of your project here...', value=topic)
    
    topic_dataframe.at[folder_name, 'Title'] = project_title
    # st.dataframe(topic_dataframe)
    st.button('Save', on_click=set_with_dataframe(get_database(parent_folder_id, 'Topic List', 'Sheet1'), topic_dataframe, include_index=True))
    
    st.write('Upload your file(s) below:')
    file_type = st.radio(
    "What type of file is it?",
    (None,'Thesis', 'Slide'), horizontal=True)
    
    col1, col2 = st.columns([0.5,0.5])
    
    with col1:

        file=st.file_uploader("Choose a PDF file", type=['pdf'], accept_multiple_files=False)

        if file:
            if file_type:
                file_name = f"{folder_name}-{file.name.split('.')[0]}-{file_type}.pdf"
                # try:
                upload_to_folder(folder_id, file_name, file)
                # except FileExistsError:
                #     st.warning(f"The file exists already. Either close the file or rename your file like so :red['{file.name.split('.')[0]}-v1.0.pdf']", icon="⚠️")
            else:
                st.warning('Make sure you select a file type', icon="⚠️")
                
    with col2:
        st.write('Current content of your personal folder:')
        resource = {
            "service_account": credentials,
            "id": f"{folder_id}",
            "fields": "files(name,id)",
        }
        res = getfilelist.GetFileList(resource)  # or r = getfilelist.GetFolderTree(resource)

        st.markdown(" <br/> ".join([f":green[{dic['name']}]" for dic in res['fileList'][0]['files']]))
