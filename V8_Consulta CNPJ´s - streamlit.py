import streamlit as st
import pandas as pd
import requests
import time
import re
import os

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="CNPJ Enterprise Analytics",
    page_icon="📊",
    layout="wide"
)

# Inicialização de Variáveis de Estado
if "rodando" not in st.session_state: st.session_state.rodando = False
if "logs" not in st.session_state: st.session_state.logs = []

ARQUIVO_SAIDA = "RESULTADOS_CNPJ.xlsx"

# --- FUNÇÕES AUXILIARES ---
def limpar_cnpj(cnpj):
    so_numeros = re.sub(r'\D', '', str(cnpj))
    return so_numeros.zfill(14) if len(so_numeros) > 0 else ""

def _get(dic, *chaves):
    for chave in chaves:
        if isinstance(dic, dict): dic = dic.get(chave, "")
        else: return ""
    return dic if dic != {} else ""

def extrair_dados_json(dados):
    estab = dados.get('estabelecimento', {})
    socios = " \n; ".join([f"Nome: {s.get('nome')} | Doc: {s.get('cpf_cnpj_socio')} | Qualificação: {_get(s, 'qualificacao_socio', 'descricao')}" for s in dados.get('socios', [])])
    secundarias = " \n; ".join([f"{act.get('subclasse')} - {act.get('descricao')}" for act in estab.get('atividades_secundarias', [])])
    
    linha = {
        "CNPJ Raiz": _get(dados, 'cnpj_raiz'), "Razão Social": _get(dados, 'razao_social'),
        "Capital Social": _get(dados, 'capital_social'), "Responsável Federativo": _get(dados, 'responsavel_federativo'),
        "Atualizado Em (Geral)": _get(dados, 'atualizado_em'), "Porte Descrição": _get(dados, 'porte', 'descricao'), 
        "Natureza Jurídica Descrição": _get(dados, 'natureza_juridica', 'descricao'), "Simples Nacional / MEI": _get(dados, 'simples'),
        "CNPJ Completo": _get(estab, 'cnpj'), "Situação Cadastral": _get(estab, 'situacao_cadastral'),
        "Data Situação Cadastral": _get(estab, 'data_situacao_cadastral'), "Data Início Atividade": _get(estab, 'data_inicio_atividade'),
        "Logradouro": _get(estab, 'logradouro'), "Número": _get(estab, 'numero'), "Bairro": _get(estab, 'bairro'),
        "CEP": _get(estab, 'cep'), "Telefone": _get(estab, 'telefone1'), "E-mail": _get(estab, 'email'), 
        "Atividade Principal": _get(estab, 'atividade_principal', 'descricao'), "Estado (UF)": _get(estab, 'estado', 'sigla'), 
        "Cidade": _get(estab, 'cidade', 'nome'), "Atividades Secundárias": secundarias, "Quadro de Sócios": socios
    }
    return linha

# --- INTERFACE ---
st.title("📊 Painel Avançado de Consulta - API CNPJ")
st.sidebar.title("⚙️ Configurações")

arquivo_carregado = st.sidebar.file_uploader("Suba a planilha Excel (Entrada)", type=["xlsx", "xls"])
df_entrada = None
coluna_selecionada = None

if arquivo_carregado:
    try:
        excel_file = pd.ExcelFile(arquivo_carregado)
        aba = st.sidebar.selectbox("Escolha a ABA:", excel_file.sheet_names)
        df_entrada = pd.read_excel(arquivo_carregado, sheet_name=aba)
        coluna_selecionada = st.sidebar.selectbox("Qual COLUNA possui os CNPJs?", df_entrada.columns)
    except Exception as e:
        st.error(f"Erro ao carregar arquivo: {e}")

# Lógica do Botão de Download
if os.path.exists(ARQUIVO_SAIDA):
    with open(ARQUIVO_SAIDA, "rb") as f:
        st.sidebar.download_button("📥 Baixar Resultado Atual", f, file_name=ARQUIVO_SAIDA, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")

# Botão Iniciar
if st.button("▶️ Iniciar Processamento", disabled=(df_entrada is None or st.session_state.rodando)):
    st.session_state.rodando = True
    st.rerun()

# --- LÓGICA DE PROCESSAMENTO ---
if st.session_state.rodando:
    lista_bruta = df_entrada[coluna_selecionada].dropna().tolist()
    resultados_atuais = []
    cnpjs_processados = set()

    # Tenta carregar progresso anterior
    if os.path.exists(ARQUIVO_SAIDA):
        try:
            df_existente = pd.read_excel(ARQUIVO_SAIDA)
            resultados_atuais = df_existente.to_dict('records')
            if 'CNPJ Completo' in df_existente.columns:
                cnpjs_processados.update([limpar_cnpj(c) for c in df_existente['CNPJ Completo'].dropna()])
            st.info(f"🔄 Retomando execução. {len(cnpjs_processados)} CNPJs já processados.")
        except: pass

    pendentes = list(dict.fromkeys([limpar_cnpj(x) for x in lista_bruta if len(limpar_cnpj(x)) == 14 and limpar_cnpj(x) not in cnpjs_processados]))
    total = len(pendentes)

    if total == 0:
        st.success("✅ Todos os CNPJs já foram processados!")
        st.session_state.rodando = False
    else:
        progress_bar = st.progress(0)
        status_text = st.empty()
        session = requests.Session()
        session.headers.update({"User-Agent": "Mozilla/5.0", "Accept": "application/json"})

        for i, cnpj in enumerate(pendentes):
            if not st.session_state.rodando: break
            
            sucesso = False
            for tentativa in range(3):
                try:
                    res = session.get(f"https://publica.cnpj.ws/cnpj/{cnpj}", timeout=15)
                    if res.status_code == 200:
                        resultados_atuais.append(extrair_dados_json(res.json()))
                        sucesso = True
                        break
                    elif res.status_code == 429:
                        time.sleep(20) # Aguarda Limite da API
                    else:
                        break
                except:
                    time.sleep(2)

            if i % 5 == 0: # Salva a cada 5 para garantir persistência
                pd.DataFrame(resultados_atuais).to_excel(ARQUIVO_SAIDA, index=False)
                progress_bar.progress((i + 1) / total)
                status_text.write(f"Processando: {i+1}/{total}")

            time.sleep(1.5) # Respeito ao delay da API

        pd.DataFrame(resultados_atuais).to_excel(ARQUIVO_SAIDA, index=False)
        st.success("🎉 Processamento Concluído!")
        st.session_state.rodando = False
        st.rerun()
