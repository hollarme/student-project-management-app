
import streamlit as st
import pandas as pd
import numpy as np
from pandas.errors import MergeError
import math
# import streamlit_authenticator as stauth

from googleapiclient.discovery import build
from google.oauth2 import service_account
# from oauth2client.service_account import ServiceAccountCredentials

import json
import gspread 
from gspread.exceptions import WorksheetNotFound, SpreadsheetNotFound, APIError

from st_aggrid import AgGrid, JsCode
from st_aggrid.grid_options_builder import GridOptionsBuilder

from gspread_dataframe import get_as_dataframe, set_with_dataframe

from annotated_text import annotated_text

import pyTigerGraph as tg

st.set_page_config(page_title="Capstone Project Manager", layout="wide") 

# Initialize connection.
conn = tg.TigerGraphConnection(**st.secrets["tigergraph"])
conn.apiToken = conn.getToken(st.secrets["tg_secret"])

# authToken = authToken[0]



checkbox_renderer = JsCode("""
class CheckboxRenderer{

    init(params) {
        this.params = params;

        this.eGui = document.createElement('input');
        this.eGui.type = 'checkbox';
        this.eGui.checked = params.value;

        this.checkedHandler = this.checkedHandler.bind(this);
        this.eGui.addEventListener('click', this.checkedHandler);
    }

    checkedHandler(e) {
        let checked = e.target.checked;
        let colId = this.params.column.colId;
        this.params.node.setDataValue(colId, checked);
    }

    getGui(params) {
        return this.eGui;
    }

    destroy(params) {
    this.eGui.removeEventListener('click', this.checkedHandler);
    }
}//end class
""")


authenticated_user = st.experimental_user.email
st.write(authenticated_user)

save_master_copy = True
disable_available_adviser = False
not_admin = False   

folder_id = None

course_name = 'EEE 501'

# Create a list of scope values to pass to the credentials object
scope = ['https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive']

drive_credentials = service_account.Credentials.from_service_account_info(st.secrets["ServiceAccountCredentials"], scopes=scope)


# Create a credentials object using the service account info and scope values
sheet_credentials = service_account.Credentials.from_service_account_info(st.secrets["ServiceAccountCredentialsSheet"], scopes = scope)

# @st.cache_resource
# class GoogleDriveService:
#     def __init__(self):
#         self._SCOPES=['https://www.googleapis.com/auth/drive', 'https://spreadsheets.google.com/feeds']
#         # self.ServiceAccountCredentials = st.secrets['ServiceAccountCredentials']
#         # st.write(st.secrets['ServiceAccountCredentials'])
#         # self.jsonString = json.dumps({key:value for key, value in self.ServiceAccountCredentials.items()})
        
#     def build(self):
#         # with open(data_file:="data.json", "w") as jf:
#         #     jf.write(self.jsonString)
#         # creds = ServiceAccountCredentials.from_json_keyfile_name(data_file, self._SCOPES)
#         service = build('drive', 'v3', credentials=drive_credentials)

#         return service
    
# @st.cache_resource
def getFileListFromGDrive(cred):
    selected_fields="files(id,name,webViewLink)"
    g_drive_service=build('drive', 'v3', credentials=cred)
    list_file=g_drive_service.files().list(fields=selected_fields).execute()
    return {"files":list_file.get("files")}

# @st.cache_resource
def get_database(folder_id, db_name, sheet):
        
        # Authorize the connection to Google Sheets using the credentials object
        gc = gspread.authorize(sheet_credentials)
        
        try:
            # Open the Google Sheets document with the specified name
            sh = gc.open(db_name, folder_id)
        except SpreadsheetNotFound:
            # Create the Google Sheets document with the specified name
            sh = gc.create(db_name, folder_id)
            # sh.share('eakinboboye@oauife.edu.ng', perm_type='user', role='editor')
        try:
            # Access the worksheet within the document with the specified name
            worksheet = sh.worksheet(sheet) 
        except WorksheetNotFound:
            #Create the worksheet for the user
            worksheet = sh.add_worksheet(sheet, rows=1000, cols=50)
            
        return worksheet
    
def send_email(recipient, score_sheet):
    from email.mime.text import MIMEText
    from email.mime.application import MIMEApplication
    from email.mime.multipart import MIMEMultipart
    from smtplib import SMTP
    import smtplib
    import sys


    recipients = [recipient] 
    emaillist = [elem.strip().split(',') for elem in recipients]
    msg = MIMEMultipart()
    msg['Subject'] = f"{course_name} Project Result"
    msg['From'] = 'eakinboboye@oauife.edu.ng'

    html = f"""\
            <html>
              <head></head>
              <body>
              <div>Please sir/ma, return grades for the students you supervised for "{course_name}".</div>
              <div>The table below contains the score sheet for your students. </div>
                {0}
              <div>Thank you.</div>
              </body>
            </html>
    """.format(score_sheet.to_html(index=False))

    part1 = MIMEText(html, 'html')
    msg.attach(part1)

    try:
        # """Checking for connection errors"""

        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login('eakinboboye@oauife.edu.ng',st.secrets["oaumailpassword"])
        server.sendmail(msg['From'], emaillist , msg.as_string())
        server.close()

    except Exception as e:
        st.write(f"Error for connection: {e}")
        
name_email_map = {
'Mr. Olorunniwo': 'dareniwo@oauife.edu.ng',
'Mr. Aransiola': 'aaransiola@oauife.edu.ng',
'Dr. Obayiuwana': 'obayiuwanae@oauife.edu.ng',
'Dr. Yesufu': 'tyesufu@oauife.edu.ng',
'Dr. Ariyo': 'ariyofunso@oauife.edu.ng',
'Dr. Ogunseye': 'aaogunseye@oauife.edu.ng',
'Mr. Olayiwola': 'solayiwola@oauife.edu.ng',
'Dr. Mrs. Offiong': 'fboffiong@oauife.edu.ng',
'Dr. Ayodele': 'kayodele@oauife.edu.ng',
'Dr. Akinwale': 'olawale.akinwale@oauife.edu.ng',
'Dr. Ilori': 'sojilori@oauife.edu.ng',
'Mr. Akinboboye': 'eakinboboye@oauife.edu.ng',
'Dr. Olawole': 'alex_olawole@oauife.edu.ng',
'Dr. Babalola': 'babfisayo@oauife.edu.ng'
}

# dareniwo@oauife.edu.ng
# aaransiola@oauife.edu.ng
# obayiuwanae@oauife.edu.ng
# tyesufu@oauife.edu.ng
# ariyofunso@oauife.edu.ng
# aaogunseye@oauife.edu.ng
# solayiwola@oauife.edu.ng
# fboffiong@oauife.edu.ng
# kayodele@oauife.edu.ng
# olawale.akinwale@oauife.edu.ng
# sojilori@oauife.edu.ng
# eakinboboye@oauife.edu.ng

# supervisor_worksheet = get_database("Supervisor Score Sheet", authenticated_user)

# preload_worksheet = get_database("defense_grouping_list", 'Sheet1')

setup,tab0, tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs(["Setup", "Information","Distribution", "Grouping", "Score Sheets", "Download Files", "I-defense Score Sheet", "I-supervisor Score Sheet", 'Result'])

with setup:
    
    st.header('Setup Page')
    
    session = st.selectbox(
    'What is the current session?',
    ('2021/2022', '2022/2023', '2023/2024', '2024/2025', '2025/2026', '2026/2027', '2027/2028', '2028/2029', '2029/2030'))
    #create folders (with the above session names) manually in eakinboboye@oauife.edu.ng and share with sheet-admin@examtools.iam.gserviceaccount.com
    
    semester = st.selectbox(
    'What is the current semester?',
    ('Rain', 'Harmattan'))
    
    global course_name
    course_name = st.selectbox(
    'What course is it?',
    ('EEE 501', 'EEE 502'))
    
    folder_name = f"{session}-{semester}"
    
    def create_files():
        gc = gspread.authorize(sheet_credentials)
        
        result = getFileListFromGDrive(sheet_credentials)
        try:
            global folder_id
            folder_id = [dic['id'] for dic in result.get('files') if dic['name']==folder_name][0]
            
            files_in_folder = [dic['name'] for dic in gc.list_spreadsheet_files(folder_id=folder_id)]
                        
            st.write(f'Files already in the folder: {",".join(files_in_folder)}.')
            
            for db_name in ['Defense Grouping List', 'Supervisor Score Sheet', 'Group Score Sheet']:
                if db_name in files_in_folder:
                    pass
                else:
                    gc.create(db_name, folder_id)
        except IndexError:
            print(f"Folder {folder_name} has not been created yet!")
            
    st.button('Create working files', on_click=create_files)
        
    
with tab0:
    # user_name = [name for name in adviser if authenticated_user.find(name.split(" ")[1].lower())==1][0]

    st.header(f'Project Management App')
   
    st.markdown('''
                _This tool was developed to help manage `EEE 501 & 502` essentially. The tool is to be employed by staff members only!_
                
                ''')
    
    with st.expander('GUIDELINE FOR EEE 501/502 Report and Presentation Format', expanded=False):
        st.markdown('''
                    The reports and presentations must be organised thus:
                    
                    1. `Introduction`: Clearly state the background to the problem

                    2. `Problem Statement`: What problem are you trying to solve and the significance of your solution

                    3. `Aim and specific objectives of the project`: What is the overall goal of your solution? What are the technical objectives required to achieve the solution?

                    4. `Literature Review`: What are the existing solutions to the problem and what are the limitations of these solutions? At least three (4) reviewed solutions should be explicitly presented in your project report. Existing reviewed solutions should be analyzed and summarized in tabular form on a single slide for your defense seminar listing the pros and cons. 
                    
                    5. `Methodology`: What is your solution? What are the step-by-step procedures used in the implementation of your solution? What are the fundamental engineering principles/theories used in your solution? Block Diagram(s)/Circuit Diagram(s)/Flow Charts of your solution or subsystem of your solution must be presented etc
                    
                    6. `Current Status of Project and Preliminary Results`: What is the current status of the project? Do you have preliminary results? Like simulations etc. What are the implications of your solution? 
                    
                    7. `Future Plans`: What are your future plans? Discuss in terms of:
                        - deliverables
                        - timeline
                        - budget, etc.
                    
                    ''')
    with st.expander('REQUIREMENTS', expanded=False):
        st.markdown('''
                    __Presentation slides__: students should prepare no more than fifteen (15) slides for the defense seminar
                     
                    __Endorsement of project report__: supervisors must endorse all reports and preview slides before projects can be assessed by the panel during defense seminar.
                   
                    __Report grading__: the student is required to revise the project report based on comments/suggestions/modifications by the assessors during the defense seminar. Revised project report and solution should be submitted to the supervisor for grading within one (1) week after your defense seminar.
                    ''')
        
    with st.expander('HOW TO SAVE FILES INTO THE PROVIDED GOOGLE DRIVE', expanded=False):
        st.markdown('''
                    You are going to save `thesis and slides) into the provided [Google drive](https://drive.google.com/drive/folders/1XJ63r2NSU3Bsv4pCiI8z1F7FFO0ugsXA?usp=sharing) using the naming convention below:2 files` (
                    
                    __Your thesis filed be named__xxxx/xxx-thesis.pdf: EEG/ shoul
                    
                    __Your presentation file hould be named__/xxx-slides.pdf: EEG/xxxxs
                    
                    > e.g. EEG/2006/059-thesis.pdf and EEG/2006/059-slides.pdf
                    
                    
                    __Note__: only pdf files are accepted!
                    
                    
                    Just put the 2 files into the `EEE 501/502 Project Management APP` folder. __Don't__ create a new folder please!
                    ''')

    
# @st.cache_data(ttl=300)
def load_data():
    # DATA_URL = "2021 EEE 501-502.csv"
    # return pd.read_csv(DATA_URL, nrows=1000)
    return get_as_dataframe(get_database(folder_id, "Defense Grouping List", 'Sheet1'), usecols=['Names','Reg. Number','Adviser','Option'])
    
def drop_na(data):
    return data.dropna(how='all')

def sort_data(data):
    return data.sort_values(by=['Option'])

def set_ind(sorted_data):
    # return sorted_data.set_index(['Adviser'])
    return sorted_data.reset_index(drop=True)

data = load_data()
ddata = drop_na(data)
sorted_data = sort_data(ddata)
indexed_data = set_ind(sorted_data)

with tab1:
    # ssmap_data = indexed_data.loc[:,['Names', 'Reg. Number', 'Adviser']]
    
    st.header(f'Staff-Student Distribution for {folder_name}')
    # AgGrid(indexed_data.loc[:,['Names', 'Reg. Number', 'Adviser', 'Option']].sort_values(by=['Adviser']))
    st.experimental_data_editor(indexed_data.loc[:,['Names', 'Reg. Number', 'Adviser']].sort_values(by=['Adviser']))
    
with tab2:
    
    preload = st.checkbox('Load the allocation done by the admin', value=not_admin)
    
    # grouping_table = []
    grouping_tables = pd.DataFrame({})
    
    adviser = indexed_data['Adviser'].drop_duplicates().values
    
    available_adviser = st.multiselect(
    'Select the available staff members:', adviser, default=[], disabled=disable_available_adviser)
    
    #get the unavailable members and change the availability status in the tigergraph
    @st.cache_data(ttl=120)
    def update_staff_availability(adviser, available_adviser):
        if available_adviser:
            unavailable_adviser = list(set(adviser)-set(available_adviser))
            for adviser in unavailable_adviser:
                conn.upsertVertex("Staff", adviser, {'available': False})
                
    update_staff_availability(adviser, available_adviser)
    
    number_of_groups = st.number_input('Insert the number of groups you want', min_value=1, disabled=not_admin)
    
    offset = 1
    groups = dict(zip(list(range(1,number_of_groups+1)),[available_adviser[i::number_of_groups] for i in range(number_of_groups)]))
    
    # st.write(groups)
    
    students_total_no = len(indexed_data['Names'].to_list())
    try:
        students_ratio = dict(zip(list(range(1,number_of_groups+1)),[round((len(value)/len(available_adviser)) * students_total_no) for value in groups.values()]))
    except ZeroDivisionError:
        students_ratio = dict(zip(list(range(1,number_of_groups+1)),[0 for value in groups.values()]))

    if not sum(students_ratio.values())==students_total_no:
        diff = students_total_no - sum(students_ratio.values())
        students_ratio[1] += diff
        
    indexed_data_copy = indexed_data.copy()
    
    indexed_data_copy.set_index('Reg. Number',inplace=True)

    random_state = st.slider('Vary randomness for best fit', 0, 100, 1)
    
    student_div = dict(zip(list(range(1,number_of_groups+1)),[[] for i in range(number_of_groups)]))
    for student_reg in indexed_data_copy.sample(frac = 1, replace=False, random_state=random_state).index:
        for key in range(1,number_of_groups+1):
            if pd.isna(student_reg):
                break
            try:
                if student_reg not in indexed_data_copy[eval("^".join([f"(indexed_data_copy.Adviser=='{i}')" for i in groups[key]]))].index and len(student_div[key])!=students_ratio[key]:
                    student_div[key].append(student_reg)
                    break
            except SyntaxError:
                pass
            
    if not preload:
        # adviser_group = dict([item for sublist in [list(zip(v,[k]*len(v))) for k,v in groups.items()] for item in sublist])
        
        grouping_tables = pd.concat(grouping_table:=[indexed_data_copy.loc[value].assign(Group=key).assign(Staff=", ".join(groups[key])) for key,value in student_div.items()])
        
        # grouping_tables.assign(StaffGroup=[adviser_group[adv] for adv in indexed_data_copy.Adviser if adv in available_adviser])
        
        if pd.concat([indexed_data.dropna().Names, grouping_tables.Names]).drop_duplicates(keep=False).shape[0] > 1:
            st.write('## Unallocated Students:', pd.concat([indexed_data.dropna().Names, grouping_tables.Names]).drop_duplicates(keep=False))
        elif pd.concat([indexed_data.dropna().Names, grouping_tables.Names]).drop_duplicates(keep=False).shape[0] == 1:
            reg_number = indexed_data.iloc[pd.concat([indexed_data.dropna().Names, grouping_tables.Names]).drop_duplicates(keep=False).index[0]]['Reg. Number']
            student_div[number_of_groups].append(student_div[1].pop())
            student_div[1].append(reg_number)

            grouping_tables = pd.concat(grouping_table:=[indexed_data_copy.loc[value].assign(Group=key).assign(Staff=", ".join(groups[key])) for key,value in student_div.items()])
    
    else:
        try:
            grouping_tables = get_as_dataframe(get_database(folder_id, "Defense Grouping List", 'Sheet1'),usecols=['Reg. Number','Names','Adviser','Group','Staff', 'Option']).dropna(how='all')#pd.read_csv('defense_grouping_list.csv', nrows=1000)
            grouping_tables.set_index('Reg. Number',inplace=True)
            grouping_table = [grouping_tables.loc[indices,:] for indices in grouping_tables.groupby('Group').groups.values()]
        except FileNotFoundError:
            grouping_tables = pd.DataFrame({})
        
    st.write('## Defense Grouping List')
    st.dataframe(grouping_tables)
    
    if save_master_copy and not grouping_tables.empty:
        st.button('save admin copy', on_click=set_with_dataframe(get_database(folder_id, "Defense Grouping List", 'Sheet1'), grouping_tables, include_index=True))
    
    #@st.experimental_memo
    def convert_df(df):
        # IMPORTANT: Cache the conversion to prevent computation on every rerun
        return df.to_csv().encode('utf-8')

    if not grouping_tables.empty:
        csv = convert_df(grouping_tables)

        st.download_button(
            label="Download list as CSV",
            data=csv,
            file_name='Defense Grouping List.csv',
            mime='text/csv',
        )
    
with tab3:
    interactive_tables = []
    
    try:
        for i,table in enumerate(grouping_table):
            table['Presentation(10)'] = np.nan
            table['Understanding(10)'] = np.nan
            table['Creativity(10)'] = np.nan
            table['Answer to Questions(5)'] = np.nan
            table['Appearance(3)'] = np.nan
            table['Attendance(2)'] = np.nan
            table['Total(40)'] = np.nan
            
            staff_list = table['Staff'].drop_duplicates().values
            
            interactive_tables.append(table)
                        
            table = table.drop(columns=['Staff', 'Adviser', 'Group']).reset_index(drop=False)
            # table.reset_index(drop=False, inplace=True)
            
            st.write(f'Group {i+1} Score Sheet: {staff_list[0]}')
            st.dataframe(table)
            
            csv = convert_df(table)

            st.download_button(
                label=f"Download Group {i+1} Score sheet as CSV",
                data=csv,
                file_name=f'Group {i+1} Score sheet.csv',
                mime='text/csv',
            )
    except (NameError, IndexError):
        st.write(f'### You have not created the defense grouping list yet!')
    # except IndexError:
        

with tab4:
    files_to_download = getFileListFromGDrive(drive_credentials)
    
    col1, col2, col3 = st.columns([2,4,4])
    
    with col1:
        filtar = st.selectbox(
        'Filter with groups',
        [f'Group {grp}' for grp in groups.keys()] if not preload else [f'Group {grp}' for grp in range(1,len(grouping_table)+1)]
        )
        
        
    with col2:
        opt = st.selectbox(
        f'Select a registration number:',
        [student for student in grouping_tables.index if (grouping_tables.loc[student].Group == int(filtar.split(" ")[1])).any()], key=2)
        
    with col3:
        # placeholder = st.empty()
        
        try:
            file = st.selectbox(
            f'Select the file to download:',
            [file['name'] for file in files_to_download['files'] if file['name'].startswith(opt.replace('/','_'))], key=3)#!='EEE 501/502 Project Management APP'], key=3)

            data =  [f['webViewLink'] for f in files_to_download['files'] if f['name'].startswith(file)]
            
            st.markdown(f"[Open {file.split('-')[1]} for {file.split('-')[0]}]({data[0]})", unsafe_allow_html=True)

            # if st.button(f"Open {file.split('-')[1]} for {file.split('-')[0]}"):
                # st.write(data[0])
                # webbrowser.open(data[0])
#             data_file = urlopen(data)
                
#             st.write(data_file.fp.read())
            
            # Download=st.download_button(label=f"Download {file.split('-')[1]} for {file.split('-')[0]}", data=b'', file_name=f'{file}',mime='application/pdf')
        except (AttributeError,TypeError):
            pass
        except IndexError:
            annotated_text(
                            "This ",
                            ("student", f"{opt}", "#8ef"),
                            " might ",
                            ("not", "#8bf"),
                            "have",
                            ("formatted", "#afa"),
                            "the",
                            ("files:", f"{file}", "#fea"),
                            "properly ",
                            "hence ",
                            "the submission is",
                            ("invalid!", "#8ef"),

                        )
            # placeholder.text('The student might not have formatted the files properly, hence the submission is invalid!')
#     with st.expander(f'Question Pool for {opt}'):
#         with st.form('Add Questions for a Student') as f:
#             question_table = pd.DataFrame({'Reg. Number':[opt],'Question':[''],'Remark':[False]})
#             # question_pool = get_as_dataframe(get_database("Question-Pool", st.session_state.authenticated_user),usecols=['Reg. Number','Question','Remark']).dropna(how='all')

#             gob = GridOptionsBuilder.from_dataframe(question_table)
#             gob.configure_column('Remark', editable=True, cellRenderer=checkbox_renderer)
#             gob.configure_default_column(editable=True)
#             gridoptions = gob.build()
#             AgGrid(question_table,
#                     gridOptions = gridoptions, 
#                     editable=True,
#                     allow_unsafe_jscode = True, 
#                     theme = 'balham',
#                     height = 200,
#                     fit_columns_on_grid_load = True)
#             st.form_submit_button("Save")
        
with tab5:
    grp_filter = st.selectbox(
        'Filter with groups',
        [f'Group {grp}' for grp in groups.keys()] if not preload else [f'Group {grp}' for grp in range(1,len(grouping_table)+1)], key=4
        )
    avg_tables = []
    table_to_edit = interactive_tables[int(grp_filter.split(" ")[1])-1]
    # st.write(table_to_edit)
    try:
        staff = table_to_edit['Staff'].drop_duplicates().values[0].split(',')#replace('. ', './').split(' ')
        staff.append(grp_filter)

        table_to_edit = table_to_edit.drop(columns=['Staff', 'Adviser', 'Group']).reset_index(drop=False)
        table_to_edit.drop(columns='Option', inplace=True)

        # st.write(staff)
        for name in staff:
            if not name.startswith('Group'):
                try:
                    int_table = get_as_dataframe(get_database(folder_id, "Group Score Sheet", name), usecols=table_to_edit.columns).dropna(how='all')#nrows=st.session_state['data_shape'][0], usecols=[i for i in range(st.session_state['data_shape'][1])])
                    table_to_edit = table_to_edit.merge(int_table, how='right')
                    table_to_edit = table_to_edit.replace(0, np.nan)
                    avg_tables.append(table_to_edit)
                except (MergeError, ValueError, APIError):
                    table_to_edit = interactive_tables[int(grp_filter.split(" ")[1])-1]
                    table_to_edit = table_to_edit.drop(columns=['Staff', 'Adviser', 'Group']).reset_index(drop=False)
                    table_to_edit.drop(columns='Option', inplace=True)
                    table_to_edit = table_to_edit.replace(0, np.nan)


            if name.startswith('Group'):
                try:
                    table_to_edit = pd.concat([table_to_edit[['Reg. Number','Names']], pd.concat(avg_tables).groupby(level=0).mean(numeric_only=True)], axis=1) #table_to_edit.merge(pd.concat(avg_tables).groupby(level=0).mean(), how='outer')
                    avg_tables = []
                except ValueError:
                    pass
            # Create a GridOptionsBuilder object from our DataFrame
#             gd = GridOptionsBuilder.from_dataframe(table_to_edit)

#             # Configure the default column to be editable
#             # sets the editable option to True for all columns
#             gd.configure_default_column(editable=True)

#             gridoptions = gd.build()
            with st.expander(f'Score Sheet for {name}'):
                with st.form(f'Score Sheet for {name}') as f:
                    response = st.data_editor(table_to_edit, use_container_width=True, key=name)
                    # response = AgGrid(table_to_edit,
                    #                 gridOptions = gridoptions, 
                    #                 editable=True,
                    #                 theme = 'balham',
                    #                 height = 200,
                    #                 fit_columns_on_grid_load = True)
                    st.write(" *Note: Don't forget to hit enter ↩ on new entry.*")
                    submit_defense = st.form_submit_button("Save")
                    if submit_defense:
                        dss_data = response.replace(np.nan, 0)#response['data'].replace(np.nan, 0)
                        cols = list(dss_data.columns)
                        cols.remove('Total(40)')
                        dss_data['Total(40)'] = dss_data[list(cols)].sum(axis=1,  numeric_only= True).values
                        set_with_dataframe(get_database(folder_id, "Group Score Sheet", name), dss_data)
                        if name.startswith('Group'):
                            dss_data = response.replace(np.nan, 0)#response['data'].replace(np.nan, 0)
                            conn.upsertVertexDataFrame(dss_data, vertexType='Average_Group_Score', v_id='Reg. Number', attributes={'reg_number':'Reg. Number', 'presentation':'Presentation(10)','creativity':'Creativity(10)','understanding':'Understanding(10)','appearance':'Appearance(3)','answer_to_questions':'Answer to Questions(5)','attendance':'Attendance(2)','total':'Total(40)'})
                            conn.upsertEdgeDataFrame(dss_data, sourceVertexType='Student', edgeType='has_group_score', targetVertexType='Average_Group_Score',from_id='Reg. Number', to_id='Reg. Number', attributes={})
    except (KeyError, IndexError):
        st.write('#### Make sure that the students are allocated into groups first or sign in to use the service!')

with tab6:
    adviser_name = st.selectbox(
        "Filter with supervisor's name",
        [name for name in adviser], key=5)
    
    # super_table = indexed_data_copy[indexed_data_copy.Adviser==[name for name in adviser if authenticated_user.find(name.split(" ")[1].lower())==1][0]]
    super_table = indexed_data[indexed_data.Adviser==adviser_name]
    super_table.drop(columns='Option', inplace=True)
    super_table = super_table.loc[:,['Names', 'Reg. Number']]#.assign(Title="").assign(Interaction(10)=""})/
    #.assign(Creativity(10)="").assign(State of Project(10)="").assign(Dedication(10)="")/
    #.assign(Quality of Report(20)="").assign.assign(Total(60)="")
    super_table['Title'] = "<insert Title>"
    super_table['Interaction(10)'] = np.nan
    super_table['Creativity(10)'] = np.nan
    super_table['State of Project(10)'] = np.nan
    super_table['Dedication(10)'] = np.nan
    super_table['Quality of Report(20)'] = np.nan
    super_table['Total(60)'] = np.nan#super_table[list(super_table.columns)].sum(axis=1,  numeric_only= True)
    try:
        super_table_remote = get_as_dataframe(get_database(folder_id, "Supervisor Score Sheet", adviser_name), usecols=['Names','Reg. Number', 'Title', 'Interaction(10)', 'Creativity(10)', 'State of Project(10)', 'Dedication(10)', 'Quality of Report(20)', 'Total(60)']).dropna(how='all')
        super_table = super_table.merge(super_table_remote, how='right')
    except (MergeError, ValueError):
        st.write('Cannot get saved table. Maybe no table is saved yet!')
    # Create a GridOptionsBuilder object from our DataFrame
    # gd = GridOptionsBuilder.from_dataframe(super_table)

    # Configure the default column to be editable
    # sets the editable option to True for all columns
    
    editable = True if 1 in [st.experimental_user.email.find(name.lower()) for name in adviser_name.split(" ") if not name.endswith('.')] else False
    if not editable:
        editable = st.experimental_user.email in ["eakinboboye@oauife.edu.ng",  'test@localhost.com']
        
    # gd.configure_side_bar()
    # gd.configure_default_column(groupable=True, value=True, aggFunc="sum", editable=editable)
    # gd.configure_column('Total(60)', valueGetter="Number(data['Interaction(10)']) + Number(data['Creativity(10)']) + Number(data['Dedication(10)']) + Number(data['State of Project(10)']) + Number(data['Quality of Report(20)'])", type=['numericColumn'])
    # gridoptions = gd.build()
        
    with st.form(f'Supervisor Score Sheet for {adviser_name}') as f:
        response = st.data_editor(super_table, use_container_width=True, key=adviser_name)
        # response = AgGrid(super_table,
        #                 gridOptions = gridoptions, 
        #                 editable=True,
        #                 theme = 'balham',
        #                 height = 200,
        #                 fit_columns_on_grid_load = True, 
        #                 enable_enterprise_modules=True)
        st.write(" *Note: Don't forget to hit enter ↩ on new entry.*")
        submit = st.form_submit_button("Save")#on_click=set_with_dataframe(get_database("Supervisor Score Sheet", adviser_name), response['data']))
        if submit:
            sss_data = response.replace(np.nan, 0)
            cols = list(sss_data.columns)
            cols.remove('Total(60)')
            sss_data['Total(60)'] = sss_data[cols].sum(axis=1,  numeric_only= True).values
            set_with_dataframe(get_database(folder_id, "Supervisor Score Sheet", adviser_name), sss_data)
            conn.upsertVertexDataFrame(sss_data, vertexType='Supervisor_Score', v_id='Reg. Number', attributes={'reg_number':'Reg. Number', 'interaction':'Interaction(10)','creativity':'Creativity(10)','state_of_project':'State of Project(10)','dedication':'Dedication(10)','quality_of_report':'Quality of Report(20)','total':'Total(60)','names':'Names'})
            conn.upsertEdgeDataFrame(sss_data, sourceVertexType='Student', edgeType='has_supervisor_score', targetVertexType='Supervisor_Score',from_id='Reg. Number', to_id='Reg. Number', attributes={})
    try:
        st.button('Send score sheet by mail', on_click=send_email, kwargs={'recipient': name_email_map[adviser_name], 'score_sheet': super_table.replace(0, " ")})
    except KeyError:
        pass

with tab7:
    try:
        _ = conn.runInterpretedQuery(
            """
            INTERPRET QUERY get_results() FOR GRAPH PMapp{

             SumAccum<FLOAT> @results;

             start = {Student.*};
             result = SELECT s FROM start:s - (has_group_score) - Average_Group_Score:a
                 ACCUM
                     s.@results += a.total;
             res = SELECT r FROM result:r - (has_supervisor_score) - Supervisor_Score:j
                ACCUM
                     r.@results += j.total
                POST-ACCUM
                     INSERT INTO Result VALUES (r.reg_number, r.@results, r.names),
                     INSERT INTO has_result VALUES (r.reg_number Student, r.reg_number Result);

             PRINT res;
            }
            """
            )
        result = conn.getVertexDataFrame('Result', select='reg_number, names, overall_total')
        result.drop(columns='v_id', inplace=True)
        result.set_axis(['Reg. Number', 'Names', 'Score'], axis='columns', inplace=True)
        result['Score'] = result['Score'].round(decimals = 0)
        buff = np.array(result['Score'].values)
        cond = np.isin(buff, [39, 49, 59, 69])
        np.add(buff, 1, out=buff, where=cond)
        result['Score'] = buff
        cat = pd.cut(result['Score'], bins=[0, 40, 45, 50, 60, 70, 100], include_lowest=True, right=False, labels=['F','E','D','C','B','A'])
        result['Grade'] = cat.tolist()
        result = st.experimental_data_editor(result)
        st.download_button(
            label="Download Result as CSV",
            data=result.to_csv().encode('utf-8'),
            file_name='Capstone Project Result.csv',
            mime='text/csv',
        )
    except KeyError:
        st.write("No result to process at the moment!")
