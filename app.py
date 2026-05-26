import base64
import gzip
import json
import re
import unicodedata
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

# =====================================================
# CONFIGURAÇÃO
# =====================================================

TEMPLATE_HTML = "template/dashboard_template.html"
OUTPUT_FOLDER = "output"

Path(OUTPUT_FOLDER).mkdir(exist_ok=True)


# =====================================================
# FUNÇÕES UTILITÁRIAS
# =====================================================

def limpar_texto(valor):
    if pd.isna(valor):
        return ""
    return str(valor).strip()


def normalizar_chave(valor):
    texto = limpar_texto(valor).upper()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = re.sub(r"\s+", " ", texto)
    return texto.strip()


def normalizar_coluna(valor):
    texto = normalizar_chave(valor)
    texto = texto.replace("/", " ")
    texto = re.sub(r"[^A-Z0-9 ]+", " ", texto)
    texto = re.sub(r"\s+", " ", texto)
    return texto.strip()


def obter_coluna(df, opcoes, obrigatoria=False):
    mapa = {normalizar_coluna(col): col for col in df.columns}

    for opcao in opcoes:
        chave = normalizar_coluna(opcao)
        if chave in mapa:
            return mapa[chave]

    for opcao in opcoes:
        chave = normalizar_coluna(opcao)
        for normalizada, original in mapa.items():
            if chave in normalizada or normalizada in chave:
                return original

    if obrigatoria:
        raise KeyError(
            f"Coluna obrigatória não encontrada. Esperado uma destas: {opcoes}. "
            f"Colunas encontradas: {list(df.columns)}"
        )

    return None


def obter_aba(xls, opcoes, obrigatoria=True):
    mapa = {normalizar_coluna(nome): nome for nome in xls.sheet_names}

    for opcao in opcoes:
        chave = normalizar_coluna(opcao)
        if chave in mapa:
            return pd.read_excel(xls, mapa[chave])

    for opcao in opcoes:
        chave = normalizar_coluna(opcao)
        for normalizada, original in mapa.items():
            if chave in normalizada or normalizada in chave:
                return pd.read_excel(xls, original)

    if obrigatoria:
        raise ValueError(
            f"Aba obrigatória não encontrada. Esperado uma destas: {opcoes}. "
            f"Abas encontradas: {xls.sheet_names}"
        )

    return pd.DataFrame()


def para_numero(valor):
    if pd.isna(valor):
        return 0.0

    if isinstance(valor, (int, float)):
        return float(valor)

    texto = str(valor).strip().replace("%", "").replace(",", ".")

    try:
        return float(texto)
    except ValueError:
        return 0.0


def limitar(valor, minimo=0, maximo=100):
    try:
        valor = float(valor)
    except Exception:
        valor = 0

    return round(max(minimo, min(maximo, valor)), 1)


def percentual(parte, total):
    if total <= 0:
        return None
    return round((parte / total) * 100, 1)


def data_valida(valor):
    try:
        if pd.isna(valor):
            return None
        return pd.to_datetime(valor, errors="coerce", dayfirst=True)
    except Exception:
        return None


def data_texto(valor):
    data = data_valida(valor)
    if data is None or pd.isna(data):
        return ""
    return data.strftime("%d/%m/%Y")


# =====================================================
# CARREGAMENTO DO EXCEL
# =====================================================

def carregar_excel(arquivo):
    xls = pd.ExcelFile(arquivo, engine="openpyxl")

    acoes = obter_aba(xls, ["AÇÕES", "ACOES", "AÇÕES ONA", "ACOES ONA"], obrigatoria=True)
    indicadores = obter_aba(xls, ["INDICADORES", "INDICADOR"], obrigatoria=True)
    notificacoes = obter_aba(xls, ["NOTIFICAÇÕES", "NOTIFICACOES", "NOTIFICAÇÃO", "NOTIFICACAO"], obrigatoria=True)
    organograma = obter_aba(xls, ["Organograma", "ORGANOGRAMA"], obrigatoria=False)

    for df in [acoes, indicadores, notificacoes, organograma]:
        if not df.empty:
            df.columns = df.columns.astype(str).str.strip()

    return acoes, indicadores, notificacoes, organograma


# =====================================================
# ORGANOGRAMA
# =====================================================

def montar_mapa_organograma(organograma):
    mapa = {}

    if organograma.empty:
        return mapa

    col_setor = obter_coluna(organograma, ["SETOR", "Setor", "Processo", "Processo Notificado"], False)
    col_resp = obter_coluna(organograma, ["Responsável pelo Setor", "Responsavel pelo Setor", "Responsável", "Responsavel"], False)
    col_dir = obter_coluna(organograma, ["Diretoria", "Direção", "Direcao"], False)
    col_ger = obter_coluna(organograma, ["Gerência", "Gerencia", "Coordenação", "Coordenacao", "Área", "Area"], False)

    if not col_setor or not col_resp:
        return mapa

    for _, row in organograma.iterrows():
        setor = limpar_texto(row.get(col_setor, ""))
        resp = limpar_texto(row.get(col_resp, ""))

        if not setor or not resp:
            continue

        mapa[normalizar_chave(setor)] = {
            "setor": setor,
            "responsavel": resp,
            "diretoria": limpar_texto(row.get(col_dir, "")) if col_dir else "",
            "gerencia": limpar_texto(row.get(col_ger, "")) if col_ger else ""
        }

    return mapa


def buscar_no_organograma(valor, mapa_org):
    chave = normalizar_chave(valor)

    if not chave:
        return None

    if chave in mapa_org:
        return mapa_org[chave]

    for chave_org, dados in mapa_org.items():
        if chave in chave_org or chave_org in chave:
            return dados

    return None


def detectar_diretoria(setor, diretoria_org=""):
    if diretoria_org:
        return diretoria_org

    texto = normalizar_chave(setor)

    tecnicos = [
        "UTI", "UICC", "UICM", "UIO", "CME", "CC", "CENTRO CIRURGICO",
        "CGO", "CPN", "PED", "PEDIATR", "PAM", "SADT", "FARMACIA",
        "FAR", "FIS", "NUT", "ENF", "PSICO", "AMB", "LAB"
    ]

    administrativos = [
        "ALM", "PATRIMONIO", "PAT", "SUP", "COMPRAS", "OPME",
        "MANUT", "ENG", "FACILITIES", "SHL", "SPR", "RH", "SPAD",
        "GESTAO DE PESSOAS", "FINANCEIRO", "FATURAMENTO", "TI"
    ]

    if any(t in texto for t in tecnicos):
        return "Diretoria Técnica"

    if any(t in texto for t in administrativos):
        return "Dir. Administrativa/Financeira"

    return "Diretoria Geral"


# =====================================================
# CLASSIFICAÇÕES E PESOS
# =====================================================

def classificar_prazo_acao(row, col_status, col_fim_prev, col_fim_real):
    status = limpar_texto(row.get(col_status, "")) if col_status else ""
    status_norm = normalizar_chave(status)

    fim_prev = data_valida(row.get(col_fim_prev, "")) if col_fim_prev else None
    fim_real = data_valida(row.get(col_fim_real, "")) if col_fim_real else None
    hoje = pd.Timestamp.today().normalize()

    if "CONCLU" in status_norm or "FINALIZ" in status_norm:
        if fim_prev is not None and fim_real is not None:
            return "CONCLUÍDA NO PRAZO" if fim_real <= fim_prev else "CONCLUÍDA COM ATRASO"
        return "CONCLUÍDA NO PRAZO"

    if "NAO INICI" in status_norm or "NÃO INICI" in status_norm:
        return "ATRASADA" if fim_prev is not None and fim_prev < hoje else "NÃO INICIADA"

    if "ATRAS" in status_norm:
        return "ATRASADA"

    if "ANDAMENTO" in status_norm or "EXECU" in status_norm:
        return "ATRASADA" if fim_prev is not None and fim_prev < hoje else "EM ANDAMENTO"

    if fim_prev is not None and fim_prev > hoje:
        return "FUTURA"

    if fim_prev is not None and fim_prev < hoje:
        return "ATRASADA"

    return "SEM DATA"


def peso_acao(tipo):
    texto = normalizar_chave(tipo)
    if "ONA" in texto:
        return 1.5
    if "PLANEJAMENTO" in texto or "ESTRATEG" in texto or texto == "PE":
        return 1.2
    return 1.0


def peso_indicador(nome, tipo):
    texto = normalizar_chave(f"{nome} {tipo}")
    if "RESULTADO" in texto or "ESTRATEG" in texto:
        return 1.5
    if "PROCESSO" in texto or "ANALISE" in texto:
        return 1.2
    return 1.0


def indicador_em_dia(status):
    texto = normalizar_chave(status)
    return "DENTRO" in texto or "OK" in texto or "CONFORME" in texto or "EM DIA" in texto


def notificacao_resolvida(situacao):
    texto = normalizar_chave(situacao)
    return any(t in texto for t in ["FINALIZADA", "ENCERRADA", "CONCLUIDA", "APROVADA", "NAO EFETIVA", "NÃO EFETIVA", "RESOLVIDA"])


def notificacao_em_andamento(situacao):
    texto = normalizar_chave(situacao)
    return any(t in texto for t in ["ACEITA", "ANALISADA", "REVISAO", "REVISÃO", "EM ANALISE", "EM ANÁLISE", "EM ACEITE", "EFETIVA", "CONVERTIDA", "EM ANDAMENTO", "TRATATIVA"])


def peso_notificacao(row, col_tipo, col_classif):
    tipo = limpar_texto(row.get(col_tipo, "")) if col_tipo else ""
    classif = limpar_texto(row.get(col_classif, "")) if col_classif else ""
    texto = normalizar_chave(f"{tipo} {classif}")

    peso = 1.0

    if "INCIDENTE" in texto:
        peso += 1.0

    if "CONFORMIDADE" in texto:
        peso += 0.7

    if "CATASTROFICO" in texto or "CATASTRÓFICO" in texto:
        peso += 4.0
    elif "GRAVE" in texto:
        peso += 3.0
    elif "MODERADO" in texto:
        peso += 2.0
    elif "LEVE" in texto:
        peso += 1.0
    elif "RISCO" in texto:
        peso += 1.0

    return peso


# =====================================================
# ESTRUTURA BASE
# =====================================================

def novo_responsavel(setor="", diretoria="", gerencia=""):
    return {
        "diretoria": detectar_diretoria(setor, diretoria),
        "gerencia": gerencia or "Não informado",
        "setores": [setor] if setor else [],
        "score": 0.0,
        "score_geral": 0.0,
        "a_score": None,
        "i_score": None,
        "n_score": None,
        "risco": {"nivel": "BAIXO", "tendencia": "POSITIVA"},
        "performance": {"acoes": 0, "indicadores": 0, "notificacoes": 0, "resolutividade": 100},
        "acoes": {
            "total": 0, "prog_geral": 0.0, "prog_por_tipo": {},
            "concl_prazo": 0, "concl_atraso": 0, "em_andamento": 0,
            "atrasada": 0, "nao_iniciada": 0, "futura": 0, "concluidas": 0
        },
        "indicadores": {
            "total_reg": 0, "inds_unicos": 0, "alim_ok": 0, "alim_atraso": 0,
            "alim_pct": None, "anal_ok": 0, "anal_atraso": 0, "anal_pct": None,
            "ind_ok": 0, "ind_prob": 0, "conformidade": None,
            "conformidade_pond": None, "dentro": 0, "fora": 0
        },
        "notificacoes": {
            "total": 0, "resolvidas": 0, "andamento": 0, "novas": 0,
            "tx_resolucao": None, "tx_resolucao_peso": None,
            "incidentes": 0, "nao_conf": 0, "ouvidoria": 0, "conflitos": 0,
            "tratativas": 0, "ev_cat": 0, "ev_grave": 0, "ev_mod": 0,
            "ev_leve": 0, "circ_risco": 0, "sem_dano": 0,
            "nao_tratadas_graves": 0, "score_sev": 0.0
        }
    }


def obter_resp(resultado, nome, setor="", diretoria="", gerencia=""):
    nome = limpar_texto(nome) or "NÃO DEFINIDO"

    if nome not in resultado:
        resultado[nome] = novo_responsavel(setor=setor, diretoria=diretoria, gerencia=gerencia)

    if setor and setor not in resultado[nome]["setores"]:
        resultado[nome]["setores"].append(setor)

    if diretoria and resultado[nome]["diretoria"] == "Diretoria Geral":
        resultado[nome]["diretoria"] = diretoria

    if gerencia and resultado[nome]["gerencia"] == "Não informado":
        resultado[nome]["gerencia"] = gerencia

    return resultado[nome]


# =====================================================
# PROCESSAMENTO DE AÇÕES
# =====================================================

def processar_acoes(acoes, resultado, detalhes, mapa_org):
    if acoes.empty:
        return

    col_titulo = obter_coluna(acoes, ["Título", "Titulo", "Plano de Ação", "Plano de Acao"], False)
    col_tipo = obter_coluna(acoes, ["Tipo"], False)
    col_setor = obter_coluna(acoes, ["Setor", "Processo"], False)
    col_resp = obter_coluna(acoes, ["Responsável", "Responsavel"], False)
    col_inicio_prev = obter_coluna(acoes, ["Início Previsto", "Inicio Previsto"], False)
    col_inicio_real = obter_coluna(acoes, ["Início Realizado", "Inicio Realizado"], False)
    col_fim_prev = obter_coluna(acoes, ["Fim Previsto"], False)
    col_fim_real = obter_coluna(acoes, ["Fim Realizado"], False)
    col_prog = obter_coluna(acoes, ["Progresso", "Percentual"], False)
    col_status = obter_coluna(acoes, ["Status", "Situação", "Situacao"], False)

    acumulador = {}

    for _, row in acoes.iterrows():
        setor = limpar_texto(row.get(col_setor, "")) if col_setor else "Não informado"
        responsavel = limpar_texto(row.get(col_resp, "")) if col_resp else ""
        dados_org = buscar_no_organograma(setor, mapa_org)

        if not responsavel and dados_org:
            responsavel = dados_org.get("responsavel", "")

        responsavel = responsavel or "NÃO DEFINIDO"
        setor_final = dados_org.get("setor", setor) if dados_org else setor
        dir_org = dados_org.get("diretoria", "") if dados_org else ""
        ger_org = dados_org.get("gerencia", "") if dados_org else ""

        item = obter_resp(resultado, responsavel, setor=setor_final, diretoria=dir_org, gerencia=ger_org)

        if responsavel not in acumulador:
            acumulador[responsavel] = {"peso_total": 0.0, "prog_ponderado": 0.0, "tipo_peso": {}}

        tipo = limpar_texto(row.get(col_tipo, "")) if col_tipo else ""
        peso = peso_acao(tipo)

        progresso = para_numero(row.get(col_prog, 0)) if col_prog else 0
        if progresso <= 1:
            progresso *= 100
        progresso = limitar(progresso)

        prazo = classificar_prazo_acao(row, col_status, col_fim_prev, col_fim_real)

        item["acoes"]["total"] += 1

        if prazo == "CONCLUÍDA NO PRAZO":
            item["acoes"]["concl_prazo"] += 1
            item["acoes"]["concluidas"] += 1
        elif prazo == "CONCLUÍDA COM ATRASO":
            item["acoes"]["concl_atraso"] += 1
            item["acoes"]["concluidas"] += 1
        elif prazo == "EM ANDAMENTO":
            item["acoes"]["em_andamento"] += 1
        elif prazo == "ATRASADA":
            item["acoes"]["atrasada"] += 1
        elif prazo == "NÃO INICIADA":
            item["acoes"]["nao_iniciada"] += 1
        elif prazo == "FUTURA":
            item["acoes"]["futura"] += 1

        acumulador[responsavel]["peso_total"] += peso
        acumulador[responsavel]["prog_ponderado"] += progresso * peso
        acumulador[responsavel]["tipo_peso"].setdefault(tipo or "Não informado", {"peso": 0.0, "prog": 0.0})
        acumulador[responsavel]["tipo_peso"][tipo or "Não informado"]["peso"] += peso
        acumulador[responsavel]["tipo_peso"][tipo or "Não informado"]["prog"] += progresso * peso

        detalhes.setdefault(responsavel, {"acoes": [], "indicadores": [], "notificacoes": []})
        detalhes[responsavel]["acoes"].append({
            "titulo": limpar_texto(row.get(col_titulo, "")) if col_titulo else "",
            "tipo": tipo or "Não informado",
            "fim_prev": data_texto(row.get(col_fim_prev, "")) if col_fim_prev else "",
            "fim_real": data_texto(row.get(col_fim_real, "")) if col_fim_real else "",
            "inicio_prev": data_texto(row.get(col_inicio_prev, "")) if col_inicio_prev else "",
            "inicio_real": data_texto(row.get(col_inicio_real, "")) if col_inicio_real else "",
            "prog": progresso,
            "prazo": prazo,
            "status": limpar_texto(row.get(col_status, "")) if col_status else ""
        })

    for responsavel, dados in acumulador.items():
        item = resultado[responsavel]
        prog = dados["prog_ponderado"] / dados["peso_total"] if dados["peso_total"] > 0 else 0
        total = max(item["acoes"]["total"], 1)
        penalidade = ((item["acoes"]["atrasada"] / total) * 25) + ((item["acoes"]["nao_iniciada"] / total) * 10)

        item["acoes"]["prog_geral"] = limitar(prog)
        item["a_score"] = limitar(prog - penalidade)

        prog_tipo = {}
        for tipo, dados_tipo in dados["tipo_peso"].items():
            if dados_tipo["peso"] > 0:
                prog_tipo[tipo] = limitar(dados_tipo["prog"] / dados_tipo["peso"])
        item["acoes"]["prog_por_tipo"] = prog_tipo


# =====================================================
# PROCESSAMENTO DE INDICADORES
# =====================================================

def processar_indicadores(indicadores, resultado, detalhes, mapa_org):
    if indicadores.empty:
        return

    col_nome = obter_coluna(indicadores, ["Indicadores", "Indicador", "Nome"], False)
    col_tipo = obter_coluna(indicadores, ["TiPO", "Tipo"], False)
    col_status = obter_coluna(indicadores, ["Status", "Situação", "Situacao"], False)
    col_resp = obter_coluna(indicadores, ["Responsavel", "Responsável"], False)
    col_setor = obter_coluna(indicadores, ["Setor", "Processo"], False)
    col_data = obter_coluna(indicadores, ["Data da Verificação", "Data da Verificacao", "Data"], False)
    col_serial = obter_coluna(indicadores, ["Serial", "ID"], False)

    acumulador = {}

    for _, row in indicadores.iterrows():
        setor = limpar_texto(row.get(col_setor, "")) if col_setor else "Não informado"
        responsavel = limpar_texto(row.get(col_resp, "")) if col_resp else ""
        dados_org = buscar_no_organograma(setor, mapa_org)

        if not responsavel and dados_org:
            responsavel = dados_org.get("responsavel", "")

        responsavel = responsavel or "NÃO DEFINIDO"
        setor_final = dados_org.get("setor", setor) if dados_org else setor
        dir_org = dados_org.get("diretoria", "") if dados_org else ""
        ger_org = dados_org.get("gerencia", "") if dados_org else ""

        item = obter_resp(resultado, responsavel, setor=setor_final, diretoria=dir_org, gerencia=ger_org)

        if responsavel not in acumulador:
            acumulador[responsavel] = {"peso_total": 0.0, "ok_peso": 0.0, "nomes": set(), "alim_total": 0, "alim_ok": 0, "anal_total": 0, "anal_ok": 0}

        nome = limpar_texto(row.get(col_nome, "")) if col_nome else ""
        tipo = limpar_texto(row.get(col_tipo, "")) if col_tipo else ""
        status = limpar_texto(row.get(col_status, "")) if col_status else ""
        peso = peso_indicador(nome, tipo)
        ok = indicador_em_dia(status)

        item["indicadores"]["total_reg"] += 1

        if nome:
            acumulador[responsavel]["nomes"].add(normalizar_chave(nome))

        if ok:
            item["indicadores"]["dentro"] += 1
            item["indicadores"]["ind_ok"] += 1
        else:
            item["indicadores"]["fora"] += 1
            item["indicadores"]["ind_prob"] += 1

        tipo_norm = normalizar_chave(tipo)

        if "ALIMENT" in tipo_norm:
            acumulador[responsavel]["alim_total"] += 1
            if ok:
                acumulador[responsavel]["alim_ok"] += 1
                item["indicadores"]["alim_ok"] += 1
            else:
                item["indicadores"]["alim_atraso"] += 1

        if "ANALISE" in tipo_norm:
            acumulador[responsavel]["anal_total"] += 1
            if ok:
                acumulador[responsavel]["anal_ok"] += 1
                item["indicadores"]["anal_ok"] += 1
            else:
                item["indicadores"]["anal_atraso"] += 1

        acumulador[responsavel]["peso_total"] += peso
        if ok:
            acumulador[responsavel]["ok_peso"] += peso

        detalhes.setdefault(responsavel, {"acoes": [], "indicadores": [], "notificacoes": []})
        detalhes[responsavel]["indicadores"].append({
            "nome": nome,
            "peso": peso,
            "tipo": tipo,
            "status": "Dentro do Prazo" if ok else "Fora do Prazo",
            "data": data_texto(row.get(col_data, "")) if col_data else "",
            "serial": limpar_texto(row.get(col_serial, "")) if col_serial else ""
        })

    for responsavel, dados in acumulador.items():
        item = resultado[responsavel]
        item["indicadores"]["inds_unicos"] = len(dados["nomes"])
        total_reg = item["indicadores"]["total_reg"]
        dentro = item["indicadores"]["dentro"]
        item["indicadores"]["conformidade"] = percentual(dentro, total_reg)
        item["indicadores"]["conformidade_pond"] = round((dados["ok_peso"] / dados["peso_total"]) * 100, 1) if dados["peso_total"] > 0 else None
        item["indicadores"]["alim_pct"] = percentual(dados["alim_ok"], dados["alim_total"])
        item["indicadores"]["anal_pct"] = percentual(dados["anal_ok"], dados["anal_total"])
        item["i_score"] = item["indicadores"]["conformidade_pond"]


# =====================================================
# PROCESSAMENTO DE NOTIFICAÇÕES
# =====================================================

def processar_notificacoes(notificacoes, resultado, detalhes, mapa_org):
    """
    Regra oficial:
    - Notificações são vinculadas SOMENTE pelo campo "Processo Notificado".
    - O campo "Responsável" da aba NOTIFICAÇÕES não é usado para atribuição.
    - Quando o Processo Notificado existe no Organograma, a notificação entra para o
      responsável do setor/processo mapeado no Organograma.
    - Quando o processo não existe no Organograma, a notificação fica em um grupo
      técnico chamado "PROCESSO NÃO MAPEADO: <processo>", evitando atribuição indevida.
    """

    if notificacoes.empty:
        return

    col_titulo = obter_coluna(notificacoes, ["Titulo", "Título"], False)
    col_serial = obter_coluna(notificacoes, ["Serial", "ID"], False)
    col_processo = obter_coluna(notificacoes, ["Processo Notificado"], False)
    col_situacao = obter_coluna(notificacoes, ["Situação", "Situacao", "Status"], False)
    col_tipo = obter_coluna(notificacoes, ["Tipo"], False)
    col_classif = obter_coluna(notificacoes, ["Classificação Incidente", "Classificacao Incidente", "Classificação", "Classificacao"], False)
    col_data = obter_coluna(notificacoes, ["Data", "Data da Notificação", "Data da Notificacao"], False)

    if not col_processo:
        raise KeyError(
            'A aba NOTIFICAÇÕES precisa ter a coluna "Processo Notificado" para vincular ocorrências por processo.'
        )

    acumulador = {}

    for _, row in notificacoes.iterrows():
        processo_original = limpar_texto(row.get(col_processo, ""))
        processo = processo_original or "NÃO INFORMADO"

        # PONTO-CHAVE: usa apenas Processo Notificado para encontrar o dono do processo.
        # Não usa o campo Responsável da notificação.
        dados_org = buscar_no_organograma(processo, mapa_org)

        if dados_org:
            responsavel = dados_org.get("responsavel", "") or f"PROCESSO NÃO MAPEADO: {processo}"
            setor = dados_org.get("setor", processo) or processo
            dir_org = dados_org.get("diretoria", "")
            ger_org = dados_org.get("gerencia", "")
        else:
            responsavel = f"PROCESSO NÃO MAPEADO: {processo}"
            setor = processo
            dir_org = "Diretoria Geral"
            ger_org = "Processo não encontrado no Organograma"

        item = obter_resp(resultado, responsavel, setor=setor, diretoria=dir_org, gerencia=ger_org)

        if responsavel not in acumulador:
            acumulador[responsavel] = {"peso_total": 0.0, "peso_resolvido": 0.0, "score_sev": 0.0}

        situacao = limpar_texto(row.get(col_situacao, "")) if col_situacao else ""
        tipo = limpar_texto(row.get(col_tipo, "")) if col_tipo else ""
        classif = limpar_texto(row.get(col_classif, "")) if col_classif else ""

        resolvida = notificacao_resolvida(situacao)
        andamento = notificacao_em_andamento(situacao)
        peso = peso_notificacao(row, col_tipo, col_classif)

        item["notificacoes"]["total"] += 1
        acumulador[responsavel]["peso_total"] += peso
        acumulador[responsavel]["score_sev"] += peso

        if resolvida:
            item["notificacoes"]["resolvidas"] += 1
            acumulador[responsavel]["peso_resolvido"] += peso
        elif andamento:
            item["notificacoes"]["andamento"] += 1
        else:
            item["notificacoes"]["novas"] += 1

        texto_tipo = normalizar_chave(tipo)
        texto_classif = normalizar_chave(classif)

        if "INCIDENTE" in texto_tipo:
            item["notificacoes"]["incidentes"] += 1
        if "CONFORMIDADE" in texto_tipo:
            item["notificacoes"]["nao_conf"] += 1
        if "OUVIDORIA" in texto_tipo:
            item["notificacoes"]["ouvidoria"] += 1
        if "CONFLITO" in texto_tipo:
            item["notificacoes"]["conflitos"] += 1
        if "TRATATIVA" in texto_tipo:
            item["notificacoes"]["tratativas"] += 1

        if "CATASTROFICO" in texto_classif:
            item["notificacoes"]["ev_cat"] += 1
            if not resolvida:
                item["notificacoes"]["nao_tratadas_graves"] += 1
        elif "GRAVE" in texto_classif:
            item["notificacoes"]["ev_grave"] += 1
            if not resolvida:
                item["notificacoes"]["nao_tratadas_graves"] += 1
        elif "MODERADO" in texto_classif:
            item["notificacoes"]["ev_mod"] += 1
        elif "LEVE" in texto_classif:
            item["notificacoes"]["ev_leve"] += 1
        elif "RISCO" in texto_classif:
            item["notificacoes"]["circ_risco"] += 1
        elif "SEM DANO" in texto_classif:
            item["notificacoes"]["sem_dano"] += 1

        detalhes.setdefault(responsavel, {"acoes": [], "indicadores": [], "notificacoes": []})
        detalhes[responsavel]["notificacoes"].append({
            "titulo": limpar_texto(row.get(col_titulo, "")) if col_titulo else "",
            "serial": limpar_texto(row.get(col_serial, "")) if col_serial else "",
            "tipo": tipo,
            "situacao": situacao,
            "classif": classif,
            "proc": processo,
            "data": data_texto(row.get(col_data, "")) if col_data else ""
        })

    for responsavel, dados in acumulador.items():
        item = resultado[responsavel]
        total = item["notificacoes"]["total"]
        resolvidas = item["notificacoes"]["resolvidas"]
        item["notificacoes"]["tx_resolucao"] = percentual(resolvidas, total)

        if dados["peso_total"] > 0:
            item["notificacoes"]["tx_resolucao_peso"] = round((dados["peso_resolvido"] / dados["peso_total"]) * 100, 1)
            item["notificacoes"]["score_sev"] = round(dados["score_sev"] / max(total, 1), 1)
        else:
            item["notificacoes"]["tx_resolucao_peso"] = None
            item["notificacoes"]["score_sev"] = 0.0

        item["n_score"] = item["notificacoes"]["tx_resolucao_peso"]


# =====================================================
# SCORE FINAL E DIRETORIAS
# =====================================================

def calcular_score_final(resultado):
    for item in resultado.values():
        componentes = []

        if item["a_score"] is not None:
            componentes.append((item["a_score"], 0.45))
        if item["i_score"] is not None:
            componentes.append((item["i_score"], 0.35))
        if item["n_score"] is not None:
            componentes.append((item["n_score"], 0.20))

        if componentes:
            soma_pesos = sum(peso for _, peso in componentes)
            score = sum(valor * peso for valor, peso in componentes) / soma_pesos
        else:
            score = 0.0

        score -= item["notificacoes"]["nao_tratadas_graves"] * 5
        item["score"] = limitar(score)
        item["score_geral"] = item["score"]

        item["performance"] = {
            "acoes": item["a_score"] if item["a_score"] is not None else 0,
            "indicadores": item["i_score"] if item["i_score"] is not None else 0,
            "notificacoes": item["n_score"] if item["n_score"] is not None else 0,
            "resolutividade": item["notificacoes"]["tx_resolucao"] if item["notificacoes"]["tx_resolucao"] is not None else 100
        }

        item["risco"] = {
            "nivel": "BAIXO" if item["score"] >= 80 else "MODERADO" if item["score"] >= 60 else "ALTO",
            "tendencia": "POSITIVA" if item["score"] >= 80 else "ATENÇÃO" if item["score"] >= 60 else "CRÍTICA"
        }


def montar_diretorias(resultado):
    diretorias = {}

    for dados in resultado.values():
        diretoria = dados.get("diretoria") or "Diretoria Geral"

        if diretoria not in diretorias:
            diretorias[diretoria] = {
                "responsaveis": 0,
                "team_score": 0.0,
                "acoes_total": 0,
                "acoes_prog": 0.0,
                "acoes_concl_prazo": 0,
                "acoes_concl_atraso": 0,
                "acoes_andamento": 0,
                "acoes_atrasadas": 0,
                "acoes_nao_ini": 0,
                "acoes_futura": 0,
                "ind": {
                    "total_reg": 0,
                    "inds_unicos": 0,
                    "ind_ok": 0,
                    "ind_prob": 0,
                    "conformidade": None,
                    "conformidade_pond": None,
                    "alim_pct": None,
                    "anal_pct": None
                },
                "notif": {
                    "total": 0,
                    "resolvidas": 0,
                    "andamento": 0,
                    "novas": 0,
                    "tx_resolucao": None,
                    "tx_resolucao_peso": None,
                    "nao_tratadas_graves": 0,
                    "score_sev": 0.0
                },
                "_score_soma": 0.0,
                "_a_soma": 0.0,
                "_a_qtd": 0,
                "_i_soma": 0.0,
                "_i_qtd": 0,
                "_n_soma": 0.0,
                "_n_qtd": 0
            }

        d = diretorias[diretoria]
        a = dados["acoes"]
        i = dados["indicadores"]
        n = dados["notificacoes"]

        d["responsaveis"] += 1
        d["_score_soma"] += dados["score"]
        d["acoes_total"] += a["total"]
        d["acoes_concl_prazo"] += a["concl_prazo"]
        d["acoes_concl_atraso"] += a["concl_atraso"]
        d["acoes_andamento"] += a["em_andamento"]
        d["acoes_atrasadas"] += a["atrasada"]
        d["acoes_nao_ini"] += a["nao_iniciada"]
        d["acoes_futura"] += a["futura"]

        if dados["a_score"] is not None:
            d["_a_soma"] += dados["a_score"]
            d["_a_qtd"] += 1

        d["ind"]["total_reg"] += i["total_reg"]
        d["ind"]["inds_unicos"] += i["inds_unicos"]
        d["ind"]["ind_ok"] += i["ind_ok"]
        d["ind"]["ind_prob"] += i["ind_prob"]

        if i["conformidade_pond"] is not None:
            d["_i_soma"] += i["conformidade_pond"]
            d["_i_qtd"] += 1

        d["notif"]["total"] += n["total"]
        d["notif"]["resolvidas"] += n["resolvidas"]
        d["notif"]["andamento"] += n["andamento"]
        d["notif"]["novas"] += n["novas"]
        d["notif"]["nao_tratadas_graves"] += n["nao_tratadas_graves"]
        d["notif"]["score_sev"] += n["score_sev"]

        if n["tx_resolucao_peso"] is not None:
            d["_n_soma"] += n["tx_resolucao_peso"]
            d["_n_qtd"] += 1

    for d in diretorias.values():
        resp = max(d["responsaveis"], 1)
        d["team_score"] = round(d["_score_soma"] / resp, 1)
        d["acoes_prog"] = round(d["_a_soma"] / d["_a_qtd"], 1) if d["_a_qtd"] else 0.0
        d["ind"]["conformidade"] = percentual(d["ind"]["ind_ok"], d["ind"]["ind_ok"] + d["ind"]["ind_prob"])
        d["ind"]["conformidade_pond"] = round(d["_i_soma"] / d["_i_qtd"], 1) if d["_i_qtd"] else None
        d["notif"]["tx_resolucao"] = percentual(d["notif"]["resolvidas"], d["notif"]["total"])
        d["notif"]["tx_resolucao_peso"] = round(d["_n_soma"] / d["_n_qtd"], 1) if d["_n_qtd"] else None

        for chave in ["_score_soma", "_a_soma", "_a_qtd", "_i_soma", "_i_qtd", "_n_soma", "_n_qtd"]:
            d.pop(chave, None)

    return diretorias


# =====================================================
# PROCESSAMENTO PRINCIPAL
# =====================================================

def processar_dados(acoes, indicadores, notificacoes, organograma):
    mapa_org = montar_mapa_organograma(organograma)
    resultado = {}
    detalhes = {}

    processar_acoes(acoes, resultado, detalhes, mapa_org)
    processar_indicadores(indicadores, resultado, detalhes, mapa_org)
    processar_notificacoes(notificacoes, resultado, detalhes, mapa_org)
    calcular_score_final(resultado)
    diretorias = montar_diretorias(resultado)

    return resultado, diretorias, detalhes


# =====================================================
# HTML
# =====================================================

def compactar_detalhes(detalhes):
    bruto = json.dumps(detalhes, ensure_ascii=False).encode("utf-8")
    return base64.b64encode(gzip.compress(bruto)).decode("utf-8")


def atualizar_html(template_html, resp_json, dir_json, detalhes_json):
    caminho_template = Path(template_html)

    if not caminho_template.exists():
        raise FileNotFoundError(f"Template HTML não encontrado: {template_html}")

    html = caminho_template.read_text(encoding="utf-8")

    resp_texto = json.dumps(resp_json, ensure_ascii=False)
    dir_texto = json.dumps(dir_json, ensure_ascii=False)
    det_b64 = compactar_detalhes(detalhes_json)

    bloco_dados = f'''<script id="data-block">
const RESP = {resp_texto};
const DIR  = {dir_texto};
</script>'''

    html, total = re.subn(
        r'<script\s+id=["\']data-block["\'][^>]*>.*?</script>',
        bloco_dados,
        html,
        flags=re.DOTALL | re.IGNORECASE
    )

    if total == 0:
        html, total_resp = re.subn(
            r'const\s+RESP\s*=\s*.*?;',
            f'const RESP = {resp_texto};',
            html,
            count=1,
            flags=re.DOTALL
        )
        html, _ = re.subn(
            r'const\s+DIR\s*=\s*.*?;',
            f'const DIR = {dir_texto};',
            html,
            count=1,
            flags=re.DOTALL
        )
        total = total_resp

    html, total_det = re.subn(
        r"const\s+DET_B64\s*=\s*['\"].*?['\"]\s*;",
        f"const DET_B64 = '{det_b64}';",
        html,
        count=1,
        flags=re.DOTALL
    )

    if total == 0:
        raise ValueError('Não foi encontrado o bloco <script id="data-block"> ou const RESP no template HTML.')

    if total_det == 0:
        html = html.replace("</body>", f"<script>const DET_B64 = '{det_b64}';</script>\n</body>")

    return html


# =====================================================
# STREAMLIT
# =====================================================

st.set_page_config(page_title="HGP Performance Center", layout="wide")

st.title("🏆 HGP Performance Center")
st.write("Upload da planilha Excel para geração automática do dashboard institucional.")

arquivo = st.file_uploader("Selecione a planilha Excel", type=["xlsx"])

if arquivo:
    try:
        with st.spinner("Processando dashboard..."):
            acoes, indicadores, notificacoes, organograma = carregar_excel(arquivo)
            dados, diretorias, detalhes = processar_dados(acoes, indicadores, notificacoes, organograma)
            html_final = atualizar_html(TEMPLATE_HTML, dados, diretorias, detalhes)

            nome_saida = f"dashboard_{datetime.now().strftime('%Y_%m_%d_%H_%M')}.html"
            caminho_saida = Path(OUTPUT_FOLDER) / nome_saida
            caminho_saida.write_text(html_final, encoding="utf-8")

        st.success("Dashboard gerado com sucesso.")

        st.download_button(
            label="⬇️ Baixar Dashboard HTML",
            data=html_final,
            file_name=nome_saida,
            mime="text/html"
        )

        st.subheader("Resumo da Geração")

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Responsáveis", len(dados))
        col2.metric("Ações", len(acoes))
        col3.metric("Indicadores", len(indicadores))
        col4.metric("Notificações", len(notificacoes))

        with st.expander("Verificação técnica"):
            st.write("Abas processadas com sucesso.")
            st.write("Primeiros responsáveis gerados:")
            st.json(list(dados.keys())[:10])

    except Exception as erro:
        st.error("Erro ao processar o dashboard.")
        st.exception(erro)

