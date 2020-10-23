import streamlit as st
import os
import pandas as pd
from docxtpl import DocxTemplate
from docx2pdf import convert
import requests,time,pythoncom
proforma_temp = os.listdir("C:/Users/Danyal/Desktop/Garibsons Setup for server/Garibson New Version 2.0 -Streamlit/Proforma Template/")
#Title

st.beta_set_page_config(page_title=st.write("Garibsons Pvt. ltd."), layout = 'wide', initial_sidebar_state = 'collapsed') #page_icon = favicon,


# hide_streamlit_style = """
#             <title>Garibsons Pvt. Ltd.</title>
#             """
# st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# hide_menu_style = """
#         <style>
#         #MainMenu {visibility: hidden;}
#         </style>
#         """
# st.markdown(hide_menu_style, unsafe_allow_html=True)
st.title("Garibsons Pvt. Ltd.")
st.text("Version 2.0")
# Defining Directories
CLIENT_FOLDER = 'C:/Users/Danyal/Desktop/Garibsons Setup for server/Garibson Web App/'
file_list = os.listdir(CLIENT_FOLDER)
dic = {key: time.ctime(os.path.getmtime(
    os.path.join(CLIENT_FOLDER, key))) for key in file_list}
page = st.selectbox("Choose the Form Type", ["Proforma Invoice","Invoice- Other Documents"])

#Selection of Page
if page == 'Proforma Invoice':
    st.header("Proforma Invoice")
    #Input from user;
    side_file = st.sidebar.radio("Please Select",proforma_temp)
    usr_input = st.text_input(label="Proforma Invoice No.", max_chars=20)
    file_name = st.file_uploader(label="Please upload Docx Template")
    submit = st.button(label="Submit")
    if submit:
        if side_file:
            file = os.path.join("C:/Users/Danyal/Desktop/Garibsons Setup for server/Garibson New Version 2.0 -Streamlit/Proforma Template/",side_file)
            try:
                doc = DocxTemplate(file)
                x = requests.get('http://151.80.237.86:1251/ords/zkt/pi_doc/doc?pi_no={}'.format(usr_input))
                data = x.json()
                for x in data['items']:
                    doc.render(x)
                    doc.save(f"{usr_input}.docx")
                    # st.success('Your file has been created!')
                    pythoncom.CoInitialize()
                    convert(f"{usr_input}.docx",f"C:/Users/Danyal/Desktop/Garibsons Setup for server/Garibson New Version 2.0 -Streamlit/{usr_input}.pdf")
                    # flash("Your file has been created!")
                st.success('Your file has been created!')
            except Exception as e:
                st.warning(e)
# Selection of other page
if page == 'Invoice- Other Documents':
    st.header("Invoice - Other Documents")
    #Input from user
    for file in os.listdir("C:/Users/Danyal/Desktop/Garibsons Setup for server/Garibson New Version 2.0 -Streamlit/Other Documents/"):
        side_file = st.sidebar.checkbox(file)
    usr_input = st.text_input(label="Invoice No.", max_chars=10)
    variable_filter = st.selectbox("Variable Filter",["Lists"] ,0)

