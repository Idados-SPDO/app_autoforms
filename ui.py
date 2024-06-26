
# Importando bibliotecas necessárias
import os
import pandas as pd
import streamlit as st
from io import BytesIO
import data_processing as dp
import tempfile
import forms_abertura as abert
import forms_ampliacao as ampli

def page_gera_forms():
    st.title("App para automatização de formulários de abertura e ampliação")

    st.write("Os dados de entrada para o app devem ser preenchidos de acordo com o arquivo abaixo.")
    modelo = pd.read_excel('TEMPLATE.xlsx')
    dp.baixar_modelo(modelo, "TEMPLATE")

    st.markdown('---')
    st.write("Importe aqui o arquivo de Input.")
    st.file_uploader("a", type="xlsx", key="content_file", label_visibility="hidden")

    st.markdown('---')
    if st.button('Gerar formulários'):
        content = dp.load_data(st.session_state.content_file)
        with tempfile.TemporaryDirectory() as temp_dir:
            abert.form_abertura(content, temp_dir)
            ampli.form_ampliacao(content, temp_dir)

            # Chama a função para compactar os arquivos e obter os dados ZIP
            zip_data = dp.zip_output_files(temp_dir)
    
            # Disponibiliza o arquivo ZIP para download
            st.download_button(
                label="Exportar arquivo",
                data=zip_data,
                file_name='Abertura_Ampliacao.zip',
                mime='application/zip'
            )
