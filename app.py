import streamlit as st
import pandas as pd
import os
import openpyxl
from io import BytesIO
from streamlit_option_menu import option_menu

# ==========================================
# CONFIGURAÇÃO DA PÁGINA
# ==========================================
st.set_page_config(
    page_title="Demand Planning APP",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==========================================
# ESTILIZAÇÃO VISUAL AVANÇADA (CSS)
# ==========================================
def aplicar_estilo_corporativo():
    estilo_css = """
    <style>
        [data-testid="stSidebar"] {
            background-color: #FFFFFF !important;
            border-right: 2px solid #4F4F4F !important;
        }
        [data-testid="stSidebar"] > div:first-child {
            background-color: #FFFFFF !important;
        }
    </style>
    """
    st.markdown(estilo_css, unsafe_allow_html=True)

aplicar_estilo_corporativo()

# ==========================================
# O ROBÔ DE PROCESSAMENTO (PANDAS + OPENPYXL)
# ==========================================
def processar_cruzamento_dr_mt(arquivo_dr, arquivo_mt, anos_selecionados, mercados_selecionados, marcas_selecionadas, mapeamento_abas, limite_mes):
    """
    Lê os arquivos, cruza a chave (Mercado, Marca, Série) e injeta os dados.
    Usa a variável 'limite_mes' para não sobrescrever fórmulas de estoque nos meses futuros.
    """
    wb_dr = openpyxl.load_workbook(arquivo_dr, data_only=True)
    wb_mt = openpyxl.load_workbook(arquivo_mt)
    
    for ano in anos_selecionados:
        aba_dr_nome = mapeamento_abas[ano]
        aba_dr = wb_dr[aba_dr_nome]
        aba_mt = wb_mt[ano]
        
        def achar_colunas(aba):
            mapa_cols = {}
            for col in range(1, 50):
                valor = aba.cell(row=4, column=col).value
                if isinstance(valor, str):
                    texto = valor.strip().upper()
                    if texto in ["MER", "MERCADO"]: mapa_cols['MERCADO'] = col
                    elif texto in ["MAR", "MARCA"]: mapa_cols['MARCA'] = col
                    elif texto in ["SÉRIE", "SERIE"]: mapa_cols['SERIE'] = col
                    elif texto in ["PRODUTO"]: mapa_cols['PRODUTO'] = col
            return mapa_cols

        cols_dr = achar_colunas(aba_dr)
        cols_mt = achar_colunas(aba_mt)
        
        if not all(k in cols_dr for k in ['MERCADO', 'MARCA', 'SERIE']):
            raise ValueError(f"Cabeçalhos não encontrados na linha 4 do arquivo DR ({aba_dr_nome}).")
            
        dados_extraidos = {}
        
        for linha in range(5, aba_dr.max_row + 1):
            mercado = aba_dr.cell(row=linha, column=cols_dr['MERCADO']).value
            marca = aba_dr.cell(row=linha, column=cols_dr['MARCA']).value
            serie = aba_dr.cell(row=linha, column=cols_dr['SERIE']).value
            produto = aba_dr.cell(row=linha, column=cols_dr.get('PRODUTO', 0)).value
            
            if not serie or not mercado or not marca:
                continue
            if isinstance(produto, str) and "TOTAL" in produto.upper():
                continue
                
            if mercado in mercados_selecionados and marca in marcas_selecionadas:
                chave = (str(mercado).strip(), str(marca).strip(), str(serie).strip())
                
                # Coleta RT (12 meses totais - Exemplo)
                valores_rt = []
                for offset in range(12):
                    valor_celula = aba_dr.cell(row=linha, column=16 + offset).value
                    valores_rt.append(valor_celula if valor_celula is not None else 0)
                
                # NOTA DO ARQUITETO: Quando formos mapear o Estoque, faremos assim:
                # valores_estoque = []
                # for offset in range(limite_mes):  <-- AQUI ENTRA A TRAVA INTELIGENTE!
                #     ...
                
                dados_extraidos[chave] = {
                    'RT': valores_rt
                    # 'ESTOQUE': valores_estoque (Entrará na próxima fase)
                }

        for linha in range(5, aba_mt.max_row + 1):
            mercado = aba_mt.cell(row=linha, column=cols_mt.get('MERCADO', 4)).value
            marca = aba_mt.cell(row=linha, column=cols_mt.get('MARCA', 5)).value
            serie = aba_mt.cell(row=linha, column=cols_mt.get('SERIE', 6)).value
            
            chave_mt = (str(mercado).strip() if mercado else "", 
                        str(marca).strip() if marca else "", 
                        str(serie).strip() if serie else "")
                        
            if chave_mt in dados_extraidos:
                # Injeta RT (12 meses)
                valores_rt = dados_extraidos[chave_mt]['RT']
                for offset in range(12):
                    aba_mt.cell(row=linha, column=16 + offset).value = valores_rt[offset]
                
                # Quando implementarmos o estoque, será injetado respeitando o limite_mes:
                # valores_estoque = dados_extraidos[chave_mt]['ESTOQUE']
                # for offset in range(limite_mes):
                #      aba_mt.cell(row=linha, column=COLUNA_ESTOQUE + offset).value = valores_estoque[offset]

    saida_virtual = BytesIO()
    wb_mt.save(saida_virtual)
    saida_virtual.seek(0)
    
    return saida_virtual

# ==========================================
# MÓDULOS DE RENDERIZAÇÃO DE TELAS
# ==========================================
def render_home():
    st.title("🏠 Planejamento de Demanda")
    st.subheader("Bem-vindo ao Portal de Demand Planning")
    st.markdown("---")
    st.info("💡 Selecione uma funcionalidade no menu lateral para iniciar o processamento.")

def render_atualizar_material():
    st.header(
        "Atualizar material de trabalho com informações do DR", 
        help="Geralmente utilizado no início do ciclo, após atualização dos realizados do mês anterior. Transfere os dados do DR para o material de trabalho: Ao carregar os dois arquivos, selecione as marcas e os mercados desejados, e as informações de RT, WS, MRP e estoques serão transferidas do DR para o Material de Trabalho!"
    )
    st.markdown("---")
    
    st.write("Faça o upload dos arquivos base para iniciar o cruzamento de dados.")
    
    col1, col2 = st.columns(2)
    with col1:
        arquivo_dr = st.file_uploader("📥 Upload do Arquivo DR", type=["xlsx", "xlsm"])
    with col2:
        arquivo_mt = st.file_uploader("📥 Upload do Material de Trabalho", type=["xlsx", "xlsm"])
        
    st.markdown("### 🎯 Parâmetros da Transferência")
    
    if arquivo_dr is not None and arquivo_mt is not None:
        try:
            abas_dr = pd.ExcelFile(arquivo_dr).sheet_names
            
            # FILTROS PRINCIPAIS
            col_ano, col_mercado, col_marca = st.columns(3)
            with col_ano:
                anos_selecionados = st.multiselect("Selecione os Anos", options=["2026", "2027"])
            with col_mercado:
                mercados_
