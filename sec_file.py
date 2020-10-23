import streamlit as st
import os,zipfile,shutil
from io import StringIO
from flask import send_from_directory
import pandas as pd
from docxtpl import DocxTemplate
from docx2pdf import convert
import requests,time,pythoncom
import base64

proforma_temp = os.listdir("C:/Users/Danyal/Desktop/Garibsons Setup for server/Garibson New Version 2.0 -Streamlit/Proforma Template/")
other_doc = os.listdir("C:/Users/Danyal/Desktop/Garibsons Setup for server/Garibson New Version 2.0 -Streamlit/Other Documents/")
#Title
favicon = "logo.png"
st.beta_set_page_config(page_title='Garibsons Pvt. Ltd.', layout = 'centered', page_icon = favicon, initial_sidebar_state = 'collapsed') #page_icon = favicon,
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
    st.sidebar.write("Please select Templates")
    side_file = [st.sidebar.checkbox(f, key=f) for f in proforma_temp]
    file_n = [file for file, checked in zip(proforma_temp, side_file) if checked]
    usr_input = st.text_input(label="Proforma Invoice No.", max_chars=20)
    file_name = st.file_uploader(label="Please upload Docx Template")
    submit = st.button(label="Submit")

    if submit:
        for file in file_n:
            x = os.path.join("C:/Users/Danyal/Desktop/Garibsons Setup for server/Garibson New Version 2.0 -Streamlit/Proforma Template/",file)
            try:
                doc = DocxTemplate(os.path.join("C:/Users/Danyal/Desktop/Garibsons Setup for server/Garibson New Version 2.0 -Streamlit/Proforma Template/",x))
                x = requests.get('http://151.80.237.86:1251/ords/zkt/pi_doc/doc?pi_no={}'.format(usr_input))
                data = x.json()
                for x in data['items']:
                    doc.render(x)
                    doc.save(f"C:/Users/Danyal/Desktop/Garibsons Setup for server/Garibson New Version 2.0 -Streamlit/Proforma/{usr_input}_{file}.docx")
                    pythoncom.CoInitialize()
                    convert(f"C:/Users/Danyal/Desktop/Garibsons Setup for server/Garibson New Version 2.0 -Streamlit/Proforma/{usr_input}_{file}.docx",f"C:/Users/Danyal/Desktop/Garibsons Setup for server/Garibson New Version 2.0 -Streamlit/Proforma/{usr_input}_{file}.pdf")
                st.success("Your file is ready to download")
            except Exception as e:
                st.warning(e)
        zipf = zipfile.ZipFile('download_pdf.zip', 'w', zipfile.ZIP_DEFLATED)
        for root, dirs, files in os.walk('./Proforma/'):
            for file in files:
                if file[-3:] == 'pdf':
                    zipf.write('./Proforma/' + file)
            zipf.close()
        zipf = zipfile.ZipFile('download_doc.zip', 'w', zipfile.ZIP_DEFLATED)
        for root, dirs, files in os.walk('./Proforma/'):
            for file in files:
                if file[-4:] == 'docx':
                    zipf.write('./Proforma/' + file)
            zipf.close()
            def get_binary_file_downloader_html(bin_file, file_label='File'):
                with open(bin_file, 'rb') as f:
                    data = f.read()
                bin_str = base64.b64encode(data).decode()
                href = f'<a href="data:application/octet-stream;base64,{bin_str}" download="{os.path.basename(bin_file)}"><button class="streamlit-button small-button primary-button ">{bin_file}</button></a>'
                return href
            pdf = st.markdown(get_binary_file_downloader_html('download_pdf.zip', 'download'), unsafe_allow_html=True)
            doc = st.markdown(get_binary_file_downloader_html('download_doc.zip', 'download'), unsafe_allow_html=True)
            if pdf and doc:
                for file in os.listdir("C:/Users/Danyal/Desktop/Garibsons Setup for server/Garibson New Version 2.0 -Streamlit/Proforma/"):
                    os.remove(os.path.join(f"C:/Users/Danyal/Desktop/Garibsons Setup for server/Garibson New Version 2.0 -Streamlit/Proforma/{file}"))
# Selection of other page
if page == 'Invoice- Other Documents':
    st.header("Invoice - Other Documents")
    st.sidebar.write("Please select Templates")
    side_file = [st.sidebar.checkbox(f, key=f) for f in other_doc]
    file_n = [file for file, checked in zip(other_doc, side_file) if checked]
    usr_input = st.text_input(label="Invoice No.", max_chars=20)
    file_name = st.file_uploader(label="Please upload Docx Template")
    submit = st.button(label="Submit")
    if submit:
        for file in file_n:
            x = os.path.join(
                "C:/Users/Danyal/Desktop/Garibsons Setup for server/Garibson New Version 2.0 -Streamlit/Other Documents/",
                file)
            try:
                doc = DocxTemplate(os.path.join(
                    "C:/Users/Danyal/Desktop/Garibsons Setup for server/Garibson New Version 2.0 -Streamlit/Other Documents/",
                    x))
                x = requests.get('http://151.80.237.86:1251/ords/zkt/pi_doc/doc?invno={}'.format(usr_input))
                data = x.json()
                for x in data['items']:
                    doc.render(x)
                    doc.save(
                        f"C:/Users/Danyal/Desktop/Garibsons Setup for server/Garibson New Version 2.0 -Streamlit/oth_docs/{file}.docx")
                    pythoncom.CoInitialize()
                    convert(
                        f"C:/Users/Danyal/Desktop/Garibsons Setup for server/Garibson New Version 2.0 -Streamlit/oth_docs/{file}.docx",
                        f"C:/Users/Danyal/Desktop/Garibsons Setup for server/Garibson New Version 2.0 -Streamlit/oth_docs/{file}.pdf")
                st.success("Your file is ready to download")
            except Exception as e:
                st.warning(e)

        zipf = zipfile.ZipFile('download_pdf.zip', 'w', zipfile.ZIP_DEFLATED)
        for root, dirs, files in os.walk('./oth_docs/'):
            for file in files:
                if file[-3:] == 'pdf':
                    zipf.write('./oth_docs/' + file)
            zipf.close()
        zipf = zipfile.ZipFile('download_doc.zip', 'w', zipfile.ZIP_DEFLATED)
        for root, dirs, files in os.walk('./oth_docs/'):
            for file in files:
                if file[-4:] == 'docx':
                    zipf.write('./oth_docs/' + file)
            zipf.close()
            def get_binary_file_downloader_html(bin_file, file_label='File'):
                with open(bin_file, 'rb') as f:
                    data = f.read()
                bin_str = base64.b64encode(data).decode()
                href = f'<a href="data:application/octet-stream;base64,{bin_str}" download="{os.path.basename(bin_file)}"><button class="streamlit-button small-button primary-button ">{bin_file}</button></a>'
                return href


            pdf = st.markdown(get_binary_file_downloader_html('download_pdf.zip', 'download'), unsafe_allow_html=True)
            doc = st.markdown(get_binary_file_downloader_html('download_doc.zip', 'download'), unsafe_allow_html=True)
            if pdf and doc:
                for file in os.listdir("C:/Users/Danyal/Desktop/Garibsons Setup for server/Garibson New Version 2.0 -Streamlit/oth_docs/"):
                    os.remove(os.path.join(f"C:/Users/Danyal/Desktop/Garibsons Setup for server/Garibson New Version 2.0 -Streamlit/oth_docs/{file}"))

