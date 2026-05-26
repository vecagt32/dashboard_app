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
# NORMALIZAR
# =====================================================

def normalizar(valor, minimo=0, maximo=100):

    if pd.isna(valor):
        return 0

    if valor < minimo:
        return minimo

    if valor > maximo:
        return maximo

    return round(float(valor), 2)

# =====================================================
# CARREGAR EXCEL
# =====================================================

def carregar_excel(arquivo):

    xls = pd.ExcelFile(
        arquivo,
        engine='openpyxl'
    )

    acoes = pd.read_excel(
        xls,
        'AÇÕES'
    )

    indicadores = pd.read_excel(
        xls,
        'INDICADORES'
    )

    notificacoes = pd.read_excel(
        xls,
        'NOTIFICAÇÕES'
    )

    organograma = pd.read_excel(
        xls,
        'Organograma'
    )

    # LIMPEZA

    for df in [
        acoes,
        indicadores,
        notificacoes,
        organograma
    ]:

        df.columns = (
            df.columns
            .astype(str)
            .str.strip()
        )

    return (
        acoes,
        indicadores,
        notificacoes,
        organograma
    )

# =====================================================
# CORRELACIONAR NOTIFICAÇÕES
# =====================================================

def correlacionar_notificacoes(
    notificacoes,
    organograma
):

    notif_corr = notificacoes.copy()

    mapa = {}

    for _, row in organograma.iterrows():

        setor = str(
            row.get(
                'SETOR',
                ''
            )
        ).strip()

        responsavel = str(
            row.get(
                'Responsável pelo Setor',
                ''
            )
        ).strip()

        mapa[setor] = responsavel

    notif_corr['Responsável Processo'] = (
        notif_corr['Processo Notificado']
        .astype(str)
        .str.strip()
        .map(mapa)
        .fillna('NÃO DEFINIDO')
    )

    return notif_corr

# =====================================================
# SCORE AÇÕES
# =====================================================

def calcular_score_acoes(df):

    if len(df) == 0:
        return 0

    progresso = (
        pd.to_numeric(
            df['Progresso'],
            errors='coerce'
        )
        .fillna(0)
        .mean()
    )

    concluidas = len(
        df[
            df['Status']
            .astype(str)
            .str.contains(
                'Concl',
                case=False,
                na=False
            )
        ]
    )

    atrasadas = len(
        df[
            df['Status']
            .astype(str)
            .str.contains(
                'Atras',
                case=False,
                na=False
            )
        ]
    )

    score = (
        (progresso * 100)
        +
        (concluidas * 2)
        -
        (atrasadas * 4)
    )

    return normalizar(score)

# =====================================================
# SCORE INDICADORES
# =====================================================

def calcular_score_indicadores(df):

    if len(df) == 0:
        return 0

    pontos = []

    for _, row in df.iterrows():

        status = str(
            row.get(
                'TiPO',
                ''
            )
        ).upper()

        if 'DENTRO' in status:

            pontos.append(100)

        elif 'FORA' in status:

            pontos.append(30)

        elif 'NÃO' in status:

            pontos.append(10)

        elif 'NAO' in status:

            pontos.append(10)

        else:

            pontos.append(60)

    return normalizar(
        sum(pontos) / len(pontos)
    )

# =====================================================
# SCORE NOTIFICAÇÕES
# =====================================================

def calcular_score_notificacoes(df):

    if len(df) == 0:
        return 100

    penalidade = 0

    for _, row in df.iterrows():

        tipo = str(
            row.get(
                'Tipo',
                ''
            )
        ).upper()

        situacao = str(
            row.get(
                'Situação',
                ''
            )
        ).upper()

        peso = 1

        if 'INCIDENTE' in tipo:
            peso = 10

        elif 'NÃO CONFORMIDADE' in tipo:
            peso = 7

        elif 'NAO CONFORMIDADE' in tipo:
            peso = 7

        if 'FINALIZADA' in situacao:

            penalidade += 0

        elif 'ANDAMENTO' in situacao:

            penalidade += peso * 2

        else:

            penalidade += peso * 4

    score = 100 - penalidade

    return normalizar(score)

# =====================================================
# RESOLUTIVIDADE
# =====================================================

def taxa_resolutividade(df):

    if len(df) == 0:
        return 100

    finalizadas = len(
        df[
            df['Situação']
            .astype(str)
            .str.contains(
                'Finalizada',
                case=False,
                na=False
            )
        ]
    )

    return normalizar(
        (finalizadas / len(df)) * 100
    )

# =====================================================
# RISCO
# =====================================================

def classificar_risco(score):

    if score >= 85:
        return 'BAIXO'

    if score >= 70:
        return 'MODERADO'

    if score >= 50:
        return 'ALTO'

    return 'CRÍTICO'

# =====================================================
# TENDÊNCIA
# =====================================================

def tendencia(score):

    if score >= 85:
        return 'EXCELENTE'

    if score >= 70:
        return 'EVOLUÇÃO'

    if score >= 50:
        return 'ATENÇÃO'

    return 'CRÍTICO'

# =====================================================
# PROCESSAMENTO PRINCIPAL
# =====================================================

def processar_responsaveis(
    acoes,
    indicadores,
    notificacoes,
    organograma
):

    responsaveis = {}

    todos = set()

    # =================================================
    # ORGANOGRAMA
    # =================================================

    todos.update(

        organograma[
            'Responsável pelo Setor'
        ]
        .dropna()
        .astype(str)
        .unique()
    )

    # =================================================
    # LOOP
    # =================================================

    for resp in todos:

        # =============================================
        # AÇÕES
        # =============================================

        dados_acoes = acoes[
            acoes['Responsável']
            .astype(str)
            .str.strip()
            == resp
        ]

        # =============================================
        # INDICADORES
        # =============================================

        if 'Responsavel' in indicadores.columns:

            dados_ind = indicadores[
                indicadores['Responsavel']
                .astype(str)
                .str.strip()
                == resp
            ]

        elif 'Responsável' in indicadores.columns:

            dados_ind = indicadores[
                indicadores['Responsável']
                .astype(str)
                .str.strip()
                == resp
            ]

        else:

            dados_ind = pd.DataFrame()

        # =============================================
        # NOTIFICAÇÕES
        # =============================================

        dados_not = notificacoes[
            notificacoes[
                'Responsável Processo'
            ]
            .astype(str)
            .str.strip()
            == resp
        ]

        # =============================================
        # SCORES
        # =============================================

        score_acoes = (
            calcular_score_acoes(
                dados_acoes
            )
        )

        score_indicadores = (
            calcular_score_indicadores(
                dados_ind
            )
        )

        score_notificacoes = (
            calcular_score_notificacoes(
                dados_not
            )
        )

        score_geral = (

            (score_acoes * 0.40)

            +

            (score_indicadores * 0.35)

            +

            (score_notificacoes * 0.25)

        )

        score_geral = normalizar(
            score_geral
        )

        # =============================================
        # RISCO
        # =============================================

        risco = classificar_risco(
            score_geral
        )

        trend = tendencia(
            score_geral
        )

        # =============================================
        # RESOLUTIVIDADE
        # =============================================

        resolutividade = (
            taxa_resolutividade(
                dados_not
            )
        )

        # =============================================
        # INCIDENTES
        # =============================================

        incidentes = len(

            dados_not[

                dados_not['Tipo']
                .astype(str)
                .str.contains(
                    'INCIDENTE',
                    case=False,
                    na=False
                )

            ]

        )

        # =============================================
        # NC
        # =============================================

        nc = len(

            dados_not[

                dados_not['Tipo']
                .astype(str)
                .str.contains(
                    'CONFORMIDADE',
                    case=False,
                    na=False
                )

            ]

        )

        # =============================================
        # FINAL
        # =============================================

        responsaveis[resp] = {

            'score_geral': score_geral,

            'risco': {

                'nivel': risco,

                'tendencia': trend

            },

            'performance': {

                'acoes': score_acoes,

                'indicadores': score_indicadores,

                'notificacoes': score_notificacoes,

                'resolutividade': resolutividade

            },

            'acoes': {

                'total': int(
                    len(dados_acoes)
                )

            },

            'indicadores': {

                'total': int(
                    len(dados_ind)
                )

            },

            'notificacoes': {

                'total': int(
                    len(dados_not)
                ),

                'incidentes_graves': int(
                    incidentes
                ),

                'nc_criticas': int(
                    nc
                ),

                'resolutividade': (
                    resolutividade
                )

            }

        }

    return responsaveis

# =====================================================
# HTML
# =====================================================

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
        r'const\\s+RESP\\s*=\\s*\\{.*?\\};',
        novo_bloco,
        html,
        flags=re.DOTALL
    )

    return html

# =====================================================
# STREAMLIT
# =====================================================

st.set_page_config(
    page_title='HGP Performance Center',
    layout='wide'
)

st.title(
    '🏆 HGP Performance Center'
)

st.write(
    'Upload da planilha Excel para geração automática do dashboard institucional.'
)

arquivo = st.file_uploader(
    'Selecione a planilha Excel',
    type=['xlsx']
)

if arquivo:

    with st.spinner(
        'Processando dashboard...'
    ):

        (
            acoes,
            indicadores,
            notificacoes,
            organograma
        ) = carregar_excel(
            arquivo
        )

        notificacoes = (
            correlacionar_notificacoes(
                notificacoes,
                organograma
            )
        )

        dados = (
            processar_responsaveis(
                acoes,
                indicadores,
                notificacoes,
                organograma
            )
        )

        html_final = atualizar_html(
            TEMPLATE_HTML,
            dados
        )

        nome_saida = (

            f'dashboard_'

            f'{datetime.now().strftime("%Y_%m_%d_%H_%M")}.html'

        )

        caminho_saida = (
            Path(OUTPUT_FOLDER)
            / nome_saida
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
