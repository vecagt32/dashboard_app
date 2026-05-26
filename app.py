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
OUTPUT_FOLDER = 'outputs'

Path(OUTPUT_FOLDER).mkdir(exist_ok=True)


# =====================================================
# LEITURA EXCEL
# =====================================================

def carregar_excel(arquivo):

    df = pd.read_excel(arquivo)

    # exemplo simples
    acoes = df
    indicadores = df
    notificacoes = df

    return acoes, indicadores, notificacoes


# =====================================================
# PROCESSAMENTO
# =====================================================

def processar_responsaveis(
    acoes,
    indicadores,
    notificacoes
):

    notif_abertas = len(notificacoes)

    tx_resolucao = 85

    responsaveis = {
        'andamento': int(notif_abertas),
        'tx_resolucao': tx_resolucao
    }

    return responsaveis


# =====================================================
# ATUALIZAR HTML
# =====================================================

def atualizar_html(template_html, dados_json):

    with open(template_html, 'r', encoding='utf-8') as f:

        html = f.read()

    novo_bloco = (
        f'const RESP={json.dumps(dados_json, ensure_ascii=False)};'
    )

    html = re.sub(
        r'const RESP=.*?;</script>',
        novo_bloco + '</script>',
        html,
        flags=re.DOTALL
    )

    return html


# =====================================================
# STREAMLIT UI
# =====================================================

st.set_page_config(
    page_title='Dashboard Qualidade',
    layout='wide'
)

st.title('📊 Gerador de Dashboard HTML')

st.write(
    'Faça upload da planilha Excel para gerar um novo dashboard HTML automaticamente.'
)

arquivo = st.file_uploader(
    'Selecione a planilha Excel',
    type=['xlsx']
)

if arquivo:

    with st.spinner('Processando dados...'):

        acoes, indicadores, notificacoes = carregar_excel(
            arquivo
        )

        dados = processar_responsaveis(
            acoes,
            indicadores,
            notificacoes
        )

        html_final = atualizar_html(
            TEMPLATE_HTML,
            dados
        )

        nome_saida = (
            f'dashboard_{datetime.now().strftime("%Y_%m_%d_%H_%M")}.html'
        )

        caminho_saida = (
            Path(OUTPUT_FOLDER) / nome_saida
        )

        with open(
            caminho_saida,
            'w',
            encoding='utf-8'
        ) as f:

            f.write(html_final)

    st.success('Dashboard gerado com sucesso.')

    st.download_button(
        label='⬇️ Baixar Dashboard HTML',
        data=html_final,
        file_name=nome_saida,
        mime='text/html'
    )

    st.subheader('Resumo da Geração')

    st.metric(
        'Responsáveis',
        len(dados)
    )

    st.metric(
        'Ações',
        len(acoes)
    )

    st.metric(
        'Indicadores',
        len(indicadores)
    )

    st.metric(
        'Notificações',
        len(notificacoes)
    )
