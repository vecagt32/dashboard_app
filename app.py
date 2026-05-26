import streamlit as st
import pandas as pd
import json
import re

from pathlib import Path
from datetime import datetime

# =====================================================
# CONFIG
# =====================================================

TEMPLATE_HTML = 'template/dashboard_template.html'
OUTPUT_FOLDER = 'output'

Path(OUTPUT_FOLDER).mkdir(exist_ok=True)

# =====================================================
# UTIL
# =====================================================


def normalizar(valor, minimo=0, maximo=100):

    if valor < minimo:
        return minimo

    if valor > maximo:
        return maximo

    return round(valor, 2)


# =====================================================
# PESOS
# =====================================================

PESO_ACOES = 0.35
PESO_INDICADORES = 0.40
PESO_NOTIFICACOES = 0.25


# =====================================================
# EXCEL
# =====================================================


def carregar_excel(arquivo):

    xls = pd.ExcelFile(
        arquivo,
        engine='openpyxl'
    )

    acoes = pd.read_excel(xls, 'AÇÕES')

    indicadores = pd.read_excel(xls, 'INDICADORES')

    notificacoes = pd.read_excel(xls, 'NOTIFICAÇÕES')

    try:

    organograma = pd.read_excel(
        xls,
        'ORGANOGRAMA'
    )

except:

    organograma = pd.DataFrame(
        columns=[
            'Processo',
            'Responsável'
        ]
    )

    return (
        acoes,
        indicadores,
    )
