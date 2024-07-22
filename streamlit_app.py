import streamlit as st
import pandas as pd
import os
from datetime import datetime
# from win32com.client import Dispatch  # Removido para compatibilidade
from fpdf import FPDF

# Função para verificar e instalar as dependências necessárias
def verificar_instalar_dependencias():
    import subprocess
    import sys

    try:
        import pandas
        import openpyxl
        # import win32com.client  # Removido para compatibilidade
        import xlsxwriter
        import PIL
        import fpdf
    except ImportError:
        st.write("Instalando dependências necessárias...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])
        st.write("Dependências instaladas.")

# Verificar e instalar dependências
verificar_instalar_dependencias()

# Funções para análise de Preenchimento, Tipos, Ortografia, Nomes, Intervalos, Condicionais

def AnalisePreenchimento(df_bd, df_matriz):
    resultado = []
    for coluna in range(1, len(df_matriz.columns)):
        texto = df_matriz.columns[coluna]
        regra = df_matriz.iloc[4, coluna]
        if texto in df_bd.columns:
            for linha in range(len(df_bd)):
                valor_celula = df_bd.loc[linha, texto]
                if (regra == "OBRIGATÓRIO" and pd.isna(valor_celula)) or (regra == "VAZIO" and not pd.isna(valor_celula)):
                    motivo = "Preenchimento obrigatório" if regra == "OBRIGATÓRIO" else "Sem preenchimento"
                    resultado.append({
                        'Data': pd.Timestamp.now(),
                        'Linha': linha+1,
                        'Coluna': texto,
                        'Valor': valor_celula,
                        'Analise': 'Preenchimento'
                    })
    return pd.DataFrame(resultado)

def AnalisarTipos(df_bd, df_matriz):
    resultados = []
    for coluna in range(2, len(df_matriz.columns)):
        texto = df_matriz.columns[coluna] 
        parametro = df_matriz.iloc[5, coluna]
        if texto in df_bd.columns:
            if parametro == "TEXTO ABC":
                filtro_alfabetico = df_bd[texto].apply(lambda x: str(x).replace(" ", "").isalpha())
                inconsistencias = df_bd[~filtro_alfabetico]
            elif parametro == "DATA":
                filtro_data_invalida = pd.to_datetime(df_bd[texto], errors='coerce').isna()
                inconsistencias = df_bd[filtro_data_invalida]
            elif parametro == "NUMERO":
                filtro_nao_numerico = pd.to_numeric(df_bd[texto], errors='coerce').isna()
                inconsistencias = df_bd[filtro_nao_numerico]
            else:
                inconsistencias = pd.DataFrame()
            for index, row in inconsistencias.iterrows():
                resultados.append({
                    'Data': pd.Timestamp.now(),
                    'Linha':index+1,
                    'Coluna': texto,
                    'Valor': row[texto],
                    'Analise': 'Tipo'
                })
    return pd.DataFrame(resultados)

def main_verificar_ortografia(df_matriz, df_bd):
    erros_ortograficos = []
    for coluna in df_matriz.iloc[6].dropna().index:
        if df_matriz.iloc[6, df_matriz.columns.get_loc(coluna)] == "DICIONÁRIO ORTOGRÁFICO":
            erros_ortograficos += verificar_ortografia(coluna, df_bd)
    return pd.DataFrame(erros_ortograficos)

def verificar_ortografia(nome_coluna, df_bd):
    # Verificação de ortografia desativada para compatibilidade
    # WordApp = iniciar_aplicativo_word()
    erros_ortograficos = []
    try:
        for indice, valor in df_bd[nome_coluna].items():
            if pd.notnull(valor):
                # if not verificar_ortografia_word(valor, WordApp):
                erros_ortograficos.append({
                    'Data': datetime.now(),
                    'Linha': indice+1,
                    'Coluna': nome_coluna,
                    'Valor': valor,
                    'Analise': 'Ortografia'
                })
    finally:
        pass
        # fechar_aplicativo_word(WordApp)
    return erros_ortograficos

# Funções de inicialização e fechamento do aplicativo Word removidas
# def iniciar_aplicativo_word():
#     WordApp = win32.Dispatch("Word.Application")
#     WordApp.Visible = False
#     return WordApp

# def fechar_aplicativo_word(WordApp):
#     if WordApp is not None:
#         WordApp.Quit()

# def verificar_ortografia_word(palavra, WordApp):
#     try:
#         return WordApp.CheckSpelling(palavra)
#     except Exception as e:
#         st.write(f"Erro ao verificar ortografia: {e}")
#         return True

def verificar_nomes(df_matriz, df_bd, df_nomes):
    erros_encontrados = []
    for coluna in df_matriz.iloc[7].dropna().index:
        if df_matriz.iloc[7, df_matriz.columns.get_loc(coluna)] == "BANCO DE NOMES":
            for indice, valor in enumerate(df_bd[coluna]):
                if pd.notnull(valor):
                    partes_nome = valor.split()
                    for parte_nome:
                        if not nome_eh_valido(parte_nome, df_nomes):
                            erros_encontrados.append({
                                'Data': datetime.now(),
                                'Linha': indice + 1,
                                'Coluna': coluna,
                                'Valor': valor,
                                'Analise': 'Nomes'
                            })
    return pd.DataFrame(erros_encontrados)

def nome_eh_valido(nome, df_nomes):
    return nome in df_nomes.iloc[:, 0].values

def verificar_intervalos(df_matriz, df_bd, df_intervalos):
    df_bd1 = df_bd.fillna(value='')
    inconsistencias = []
    row = df_matriz.iloc[8]
    for col in df_matriz.columns:
        if 'intervalo' in str(row[col]):
            intervalo = str(row[col])
            if col in df_bd1.columns:
                for idx, valor_bd in enumerate(df_bd1[col]):
                    try:
                        float(valor_bd)
                        if valor_bd not in df_intervalos[intervalo]:
                            hora_analise = datetime.now()
                            inconsistencias.append({
                                'Data': hora_analise,
                                'Linha': idx + 1,
                                'Coluna': col,
                                'Valor': valor_bd,
                                'Analise': 'Intervalos'
                            })
                    except:
                        if str(valor_bd) not in str(df_intervalos[intervalo]):
                            hora_analise = datetime.now()
                            inconsistencias.append({
                                'Data': hora_analise,
                                'Linha': idx + 1,
                                'Coluna': col,
                                'Valor': valor_bd,
                                'Analise': 'Intervalos'
                            })
    return pd.DataFrame(inconsistencias)

def verificar_condicionais(df_matriz, df_bd):
    inconsistencias = []
    df_matriz1 = df_matriz.fillna(value='')
    for linha in range(15, 92, 4):
        linha_matriz = df_matriz1.iloc[linha]
        for col in range(3, len(linha_matriz)):
            valor = linha_matriz.iloc[col]
            if valor != "":
                coluna_condicionante = valor
                coluna_condicionada = df_matriz.columns[col]
                dado_condicionante = df_matriz1.iloc[linha + 1, col]
                resultado = df_matriz1.iloc[linha + 2, col]
                if coluna_condicionante in df_bd.columns:
                    for indice, valor_bd in enumerate(df_bd[coluna_condicionante]):
                        if valor_bd == dado_condicionante:
                            if coluna_condicionada in df_bd.columns:
                                if df_bd[coluna_condicionada][indice] != resultado:
                                    inconsistencias.append({
                                        'Data': datetime.now(),
                                        'Linha': indice + 1,
                                        'Coluna': coluna_condicionada,
                                        'Valor': valor_bd,
                                        'Analise': 'Condicionais'
                                    })
                else:
                    st.write(f'Coluna condicionante "{coluna_condicionante}" não encontrada em df_bd.')
    return pd.DataFrame(inconsistencias)

# Função para selecionar arquivo de banco de dados
def selecionar_arquivo_banco_dados():
    caminho_banco_de_dados = st.file_uploader("Selecione o arquivo Excel do banco de dados", type=["xlsx"])
    if caminho_banco_de_dados:
        try:
            global df_bd
            df_bd = pd.read_excel(caminho_banco_de_dados)
            st.success("Arquivo de banco de dados carregado com sucesso.")
        except Exception as e:
            st.error(f"Ocorreu um erro ao carregar o arquivo de banco de dados: {str(e)}")

# Função para selecionar arquivo da matriz
def selecionar_arquivo_matriz():
    caminho_matriz = st.file_uploader("Selecione o arquivo Excel da matriz", type=["xlsx"])
    if caminho_matriz:
        try:
            global df_matriz
            df_matriz = pd.read_excel(caminho_matriz, header=1)
            st.success("Arquivo de matriz carregado com sucesso.")
        except Exception as e:
            st.error(f"Ocorreu um erro ao carregar o arquivo de matriz: {str(e)}")

# Função para selecionar arquivo de nomes
def selecionar_arquivo_nomes():
    caminho_nomes = st.file_uploader("Selecione o arquivo Excel de nomes", type=["xlsx"])
    if caminho_nomes:
        try:
            global df_nomes
            df_nomes = pd.read_excel(caminho_nomes)
            st.success("Arquivo de nomes carregado com sucesso.")
        except Exception as e:
            st.error(f"Ocorreu um erro ao carregar o arquivo de nomes: {str(e)}")

# Função para selecionar arquivo de intervalos
def selecionar_arquivo_intervalos():
    caminho_intervalos = st.file_uploader("Selecione o arquivo Excel de intervalos", type=["xlsx"])
    if caminho_intervalos:
        try:
            global df_intervalos
            df_intervalos = pd.read_excel(caminho_intervalos)
            st.success("Arquivo de intervalos carregado com sucesso.")
        except Exception as e:
            st.error(f"Ocorreu um erro ao carregar o arquivo de intervalos: {str(e)}")

# Função para iniciar a verificação de análise de preenchimento
def iniciar_analise_preenchimento():
    if 'df_bd' in globals() and 'df_matriz' in globals():
        resultado = AnalisePreenchimento(df_bd, df_matriz)
        st.write(resultado)
    else:
        st.warning('Selecione os arquivos de banco de dados e matriz antes de iniciar a análise.')

# Função para iniciar a verificação de análise de tipos
def iniciar_analise_tipos():
    if 'df_bd' in globals() and 'df_matriz' in globals():
        resultado = AnalisarTipos(df_bd, df_matriz)
        st.write(resultado)
    else:
        st.warning('Selecione os arquivos de banco de dados e matriz antes de iniciar a análise.')

# Função para iniciar a verificação de ortografia
def iniciar_verificacao_ortografia():
    if 'df_bd' in globals() and 'df_matriz' in globals():
        resultado = main_verificar_ortografia(df_matriz, df_bd)
        st.write(resultado)
    else:
        st.warning('Selecione os arquivos de banco de dados e matriz antes de iniciar a verificação.')

# Função para iniciar a verificação de nomes
def iniciar_verificacao_nomes():
    if 'df_bd' in globals() and 'df_matriz' in globals() and 'df_nomes' in globals():
        resultado = verificar_nomes(df_matriz, df_bd, df_nomes)
        st.write(resultado)
    else:
        st.warning('Selecione os arquivos de banco de dados, matriz e nomes antes de iniciar a verificação.')

# Função para iniciar a verificação de intervalos
def iniciar_verificacao_intervalos():
    if 'df_bd' in globals() and 'df_matriz' in globals() and 'df_intervalos' in globals():
        resultado = verificar_intervalos(df_matriz, df_bd, df_intervalos)
        st.write(resultado)
    else:
        st.warning('Selecione os arquivos de banco de dados, matriz e intervalos antes de iniciar a verificação.')

# Função para iniciar a verificação de condicionais
def iniciar_verificacao_condicionais():
    if 'df_bd' in globals() and 'df_matriz' in globals():
        resultado = verificar_condicionais(df_matriz, df_bd)
        st.write(resultado)
    else:
        st.warning('Selecione os arquivos de banco de dados e matriz antes de iniciar a verificação.')

# Função para concatenar todos os resultados em um relatório
def concatenar_resultados():
    global resultado_concatenado
    if 'df_bd' in globals() and 'df_matriz' in globals():
        resultado1 = AnalisePreenchimento(df_bd, df_matriz)
        resultado2 = AnalisarTipos(df_bd, df_matriz)
        resultado3 = main_verificar_ortografia(df_matriz, df_bd)
        resultado4 = verificar_nomes(df_matriz, df_bd, df_nomes) se 'df_nomes' in globals() else pd.DataFrame()
        resultado5 = verificar_intervalos(df_matriz, df_bd, df_intervalos) se 'df_intervalos' in globals() else pd.DataFrame()
        resultado6 = verificar_condicionais(df_matriz, df_bd)
        resultado_concatenado = pd.concat([resultado1, resultado2, resultado3, resultado4, resultado5, resultado6], ignore_index=True)
        st.write(resultado_concatenado.head(20))
    else:
        st.warning('Selecione todos os arquivos necessários antes de executar as verificações.')

# Interface do Streamlit
st.title('Verificador de Inconsistências - NMC')

# Carregar arquivos
st.sidebar.header('Importar Arquivos')
selecionar_arquivo_banco_dados()
selecionar_arquivo_matriz()
selecionar_arquivo_nomes()
selecionar_arquivo_intervalos()

# Executar análises
st.sidebar.header('Executar Análises')
if st.sidebar.button('Análise de Preenchimento'):
    iniciar_analise_preenchimento()
if st.sidebar.button('Análise de Tipos'):
    iniciar_analise_tipos()
if st.sidebar.button('Verificar Ortografia'):
    iniciar_verificacao_ortografia()
if st.sidebar.button('Verificar Nomes'):
    iniciar_verificacao_nomes()
if st.sidebar.button('Verificar Intervalos'):
    iniciar_verificacao_intervalos()
if st.sidebar.button('Verificar Condicionais'):
    iniciar_verificacao_condicionais()
if st.sidebar.button('Executar Verificações'):
    concatenar_resultados()
