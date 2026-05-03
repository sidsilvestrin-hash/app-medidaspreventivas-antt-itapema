"""Launcher principal do app completo de Medidas Preventivas ANTT.

O app completo fica em app.py. Este arquivo existe para que o comando
`streamlit run main.py` abra todas as funcionalidades: login, dashboard,
cadastro, importacao, pendencias, auditoria e administracao.
"""

import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
os.chdir(BASE_DIR)

from app import main as app_main


if __name__ == "__main__":
    app_main()
