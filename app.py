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
# FUNÇÕES
# =====================================================

def percentual(parte, total):
    if total == 0:
        return 0
    return round((parte / total) * 100, 1)


def carregar_excel(arquivo):

    xls = pd.ExcelFile(arquivo)

    acoes = pd.read_excel(xls, 'AÇÕES')
    indicadores = pd.read_excel(xls, 'INDICADORES')
    notificacoes = pd.read_excel(xls, 'NOTIFICAÇÕES')

    return acoes, indicadores, notificacoes


def detectar_diretoria(setor):

    setor = str(setor).upper()

    tecnica = [
        'UTI',
        'CC',
        'CGO',
        'PED',
        'PAM'
    ]

    adm = [
        'ALM',
        'PAT',
        'SUP'
    ]

    for item in tecnica:
        if item in setor:
            return 'Diretoria Técnica'

    for item in adm:
        if item in setor:
            return 'Diretoria Administrativa'

    return 'Diretoria Geral'


def processar_responsaveis(
    acoes,
    indicadores,
    notificacoes
):

    responsaveis = {}

    todos = set()

    todos.update(
        acoes['Responsável'].dropna().unique()
    )

    if 'Responsavel' in indicadores.columns:

        todos.update(
            indicadores['Responsavel'].dropna().unique()
        )

    todos.update(
        notificacoes['Responsável'].dropna().unique()
    )

    for resp in todos:

        dados_acoes = acoes[
            acoes['Responsável'] == resp
        ]

        dados_ind = indicadores[
            indicadores['Responsavel'] == resp
        ]

        dados_not = notificacoes[
            notificacoes['Responsável'] == resp
        ]

        setores = []

        if len(dados_acoes):

            setores += (
                dados_acoes['Setor']
                .dropna()
                .unique()
                .tolist()
            )

        if len(dados_ind):

            setores += (
                dados_ind['Setor']
                .dropna()
                .unique()
                .tolist()
            )

        setores = list(set(setores))

        setor_base = (
            setores[0]
            if setores
            else 'Não informado'
        )

        diretoria = detectar_diretoria(setor_base)

        total_acoes = len(dados_acoes)

        concluidas = len(
            dados_acoes[
                dados_acoes['Status']
                .astype(str)
                .str.contains(
                    'Concl',
                    case=False,
                    na=False
                )
            ]
        )

        andamento = len(
            dados_acoes[
                dados_acoes['Status']
                .astype(str)
                .str.contains(
                    'Andamento',
                    case=False,
                    na=False
                )
            ]
        )

        nao_iniciada = len(
            dados_acoes[
                dados_acoes['Status']
                .astype(str)
                .str.contains(
                    'Não',
                    case=False,
                    na=False
                )
            ]
        )

        progresso = 0

        if total_acoes > 0:

            progresso = round(
                dados_acoes['Progresso']
                .fillna(0)
                .mean() * 100,
                1
            )

        indicadores_ok = len(
            dados_ind[
                dados_ind['Status']
                .astype(str)
                .str.contains(
                    'Dentro',
                    case=False,
                    na=False
                )
            ]
        )

        indicadores_fora = len(
            dados_ind[
                dados_ind['Status']
                .astype(str)
                .str.contains(
                    'Fora',
                    case=False,
                    na=False
                )
            ]
        )

        total_ind = indicadores_ok + indicadores_fora

        conformidade = percentual(
            indicadores_ok,
            total_ind
        )

        notif_total = len(dados_not)

        notif_resolvidas = len(
            dados_not[
                dados_not['Situação']
                .astype(str)
                .str.contains(
                    'Finalizada',
                    case=False,
                    na=False
                )
            ]
        )

        notif_abertas = (
            notif_total - notif_resolvidas
        )

        tx_resolucao = percentual(
            notif_resolvidas,
            notif_total
        )

        responsaveis[resp] = {

            'diretoria': diretoria,

            'gerencia': 'Automática',

            'setores': setores,

            'acoes': {
                'total': int(total_acoes),
                'prog_geral': progresso,
                'concluidas': int(concluidas),
                'em_andamento': int(andamento),
                'nao_iniciada': int(nao_iniciada),
                'atrasada': 0
            },

            'indicadores': {
                'total_reg': int(total_ind),
                'dentro': int(indicadores_ok),
                'fora': int(indicadores_fora),
                'conformidade': conformidade
            },

            'notificacoes': {
                'total': int(notif_total),
                'resolvidas': int(notif_resolvidas),
                'andamento': int(notif_abertas),
                'tx_resolucao': tx_resolucao
            }
        }

    return responsaveis


def atualizar_html(
    template_html,
    dados_json
):

    with open(
        template_html,
        'r',
        encoding='utf-8'
    ) as f:

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

st.title(
    '📊 Gerador de Dashboard HTML'
)

st.write(
    'Faça upload da planilha Excel para gerar um novo dashboard HTML automaticamente.'
)

arquivo = st.file_uploader(
    'Selecione a planilha Excel',
    type=['xlsx']
)

if arquivo:

    with st.spinner('Processando dados...'):

        acoes, indicadores, notificacoes = carregar_excel(arquivo)

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

    st.success(
        'Dashboard gerado com sucesso.'
    )

    st.download_button(
        label='⬇️ Baixar Dashboard HTML',
        data=html_final,
        file_name=nome_saida,
        mime='text/html'
    )

    st.subheader(
        'Resumo da Geração'
    )

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
