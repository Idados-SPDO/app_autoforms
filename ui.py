# -*- coding: ansi -*-

# Importando bibliotecas necess�rias
import os
import pandas as pd
import streamlit as st
from io import BytesIO
import data_processing as dp

def page_gera_forms():
    st.title("App para automatiza��o de formul�rios de abertura e amplia��o")

    st.write("Os dados de entrada para o app devem ser preenchidos de acordo com o arquivo abaixo.")
    modelo = pd.read_excel('TEMPLATE.xlsx')
    dp.baixar_modelo(modelo, "TEMPLATE")

    st.markdown('---')
    st.write("Importe aqui o arquivo de Input.")
    st.file_uploader("a", type="xlsx", key="content_file", label_visibility="hidden")

    st.markdown('---')
    if st.button('Gerar formul�rios'):
        content = dp.load_data(st.session_state.content_file)
        exec(open('scr/code/01_forms_abertura.py').read())
        exec(open('scr/code/02_forms_ampliacao.py').read())

        # Chama a fun��o para compactar os arquivos e obter os dados ZIP
        pasta_arquivo = os.path.join(os.getcwd(), "output")
        zip_data = dp.zip_output_files(pasta_arquivo)
    
        # Disponibiliza o arquivo ZIP para download
        st.download_button(
            label="Exportar arquivo",
            data=zip_data,
            file_name='Abertura_Ampliacao.zip',
            mime='application/zip'
        )
