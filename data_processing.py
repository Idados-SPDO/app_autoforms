
# Importando bibliotecas necessárias
import os
import io
from io import BytesIO
import streamlit as st
import zipfile
import pandas as pd

def load_data(content_file=None):
    if content_file is not None:
        dtype = {
            'col1': 'str', 'col2': 'datetime64[ns]', 'col3': 'str', 'col4': 'float', 'col5': 'str', 'col6': 'str', 'col7': 'str', 'col8': 'str',
            'col9': 'str', 'col10': 'str', 'col11': 'str', 'col12': 'str', 'col13': 'str', 'col14': 'str', 'col15': 'str', 'col16': 'str',
            'col17': 'str', 'col18': 'str', 'col19': 'datetime64[ns]', 'col20': 'str', 'col21': 'str', 'col22': 'str', 'col23': 'datetime64[ns]', 'col24': 'str'
        }

        content = pd.read_excel(content_file, dtype=dtype)

        return content

def zip_output_files(output_folder):
    # Cria um objeto BytesIO para armazenar o arquivo ZIP em mem�ria
    zip_buffer = BytesIO()
    
    # Cria o arquivo ZIP em memória
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for foldername, subfolders, filenames in os.walk(output_folder):
            for filename in filenames:
                file_path = os.path.join(foldername, filename)
                arcname = os.path.relpath(file_path, output_folder)
                zip_file.write(file_path, arcname)
    
    # Move o ponteiro de volta ao início do buffer
    zip_buffer.seek(0)
    
    return zip_buffer

def baixar_modelo(df, arquivo):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=arquivo)
    buffer.seek(0)  # Volte ao início do buffer
    return st.download_button(label="Baixar Modelo de Input", data=buffer, file_name=f"{arquivo}.xlsx")

def temp_paste():
     return tempfile.TemporaryDirectory()
