# -*- coding: ansi -*-

# Importando bibliotecas necessárias
import pandas as pd
import numpy as np
import re
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.styles.fills import PatternFill
from openpyxl import Workbook
from openpyxl import load_workbook
import os

# Função para limpar o ambiente
def clear_environment():
    for name in dir():
        if not name.startswith('_'):
            del globals()[name]

# Limpar o ambiente
clear_environment()

# Limpar a tela do console
os.system('cls' if os.name == 'nt' else 'clear')

# Função principal para leitura e execução de scripts
def main():
    # Executando os scripts para exportar os formulários
    exec(open('scr/code/01_forms_abertura.py').read())
    exec(open('scr/code/02_forms_ampliacao.py').read())
    
    # Mensagem final
    print("\n Formulários gerados com sucesso! :)")

# Executar a função principal
if __name__ == "__main__":
    main()