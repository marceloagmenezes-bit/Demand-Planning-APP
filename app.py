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
# FUNÇÕES AUXILIARES DE LIMPEZA
# ==========================================
def limpar_texto(texto):
    if texto is None:
        return ""
    return " ".join(str(texto).upper().strip().split())

# ==========================================
# O ROBÔ DE PROCESSAMENTO (PANDAS + OPENPYXL)
# ==========================================
def processar_cruzamento_dr_mt(arquivo_dr, arquivo_mt, anos_selecionados, mercados_selecionados, marcas_selecionadas, mapeamento_abas, limite_mes):
    wb_dr = openpyxl.load_workbook(arquivo_dr, data_only=True)
    wb_mt = openpyxl.load_workbook(arquivo_mt, keep_vba=True)
    
    linhas_atualizadas_total = 0
    mercados_sel_clean = [limpar_texto(m) for m in mercados_selecionados]
    marcas_sel_clean = [limpar_texto(m) for m in marcas_selecionadas]

    for ano in anos_selecionados:
        aba_dr_nome = mapeamento_abas[ano]
        aba_dr = wb_dr[aba_dr_nome]
        
        if ano not in wb_mt.sheetnames:
            raise ValueError(f"A aba '{ano}' não foi encontrada no Material de Trabalho.")
        aba_mt = wb_mt[ano]
        
        def achar_colunas(aba):
            mapa_cols = {}
            for col in range(1, 50):
                valor = aba.cell(row=4, column=col).value
                if isinstance(valor, str):
                    texto = limpar_texto(valor)
                    if texto.startswith("MER"): mapa_cols['MERCADO'] = col
                    elif texto.startswith("MAR"): mapa_cols['MARCA'] = col
                    elif "SÉRI" in texto or "SERI" in texto: mapa_cols['SERIE'] = col
                    elif "PRODU" in texto: mapa_cols['PRODUTO'] = col
            return mapa_cols

        cols_dr = achar_colunas(aba_dr)
        cols_mt = achar_colunas(aba_mt)
        
        if not all(k in cols_dr for k in ['MERCADO', 'MARCA', 'SERIE']):
            raise ValueError(f"Cabeçalhos não encontrados na linha 4 do arquivo DR ({aba_dr_nome}).")
            
        dados_extraidos = {}
        
        vazios_consecutivos = 0
        for linha in range(5, aba_dr.max_row + 1):
            mercado_raw = aba_dr.cell(row=linha, column=cols_dr['MERCADO']).value
            marca_raw = aba_dr.cell(row=linha, column=cols_dr['MARCA']).value
            serie_raw = aba_dr.cell(row=linha, column=cols_dr['SERIE']).value
            produto_raw = aba_dr.cell(row=linha, column=cols_dr.get('PRODUTO', 0)).value
            
            if not serie_raw:
                vazios_consecutivos += 1
                if vazios_consecutivos > 15:
                    break
                continue
            
            vazios_consecutivos = 0
            mercado = limpar_texto(mercado_raw)
            marca = limpar_texto(marca_raw)
            serie = limpar_texto(serie_raw)
            produto = limpar_texto(produto_raw)
            
            if "TOTAL" in produto:
                continue
                
            if mercado in mercados_sel_clean and marca in marcas_sel_clean:
                chave = (mercado, marca, serie)
                valores_rt = []
                for offset in range(12):
                    valor_celula = aba_dr.cell(row=linha, column=16 + offset).value
                    valores_rt.append(valor_celula if valor_celula is not None else 0)
                
                dados_extraidos[chave] = {
                    'RT': valores_rt
                }

        vazios_consecutivos_mt = 0
        for linha in range(5, aba_mt.max_row + 1):
            col_mer = cols_mt.get('MERCADO', 4)
            col_mar = cols_mt.get('MARCA', 5)
            col_ser = cols_mt.get('SERIE', 7)
            
            mercado_raw = aba_mt.cell(row=linha, column=col_mer).value
            marca_raw = aba_mt.cell(row=linha, column=col_mar).value
            serie_raw = aba_mt.cell(row=linha, column=col_ser).value
            
            if not serie_raw:
                vazios_consecutivos_mt += 1
                if vazios_consecutivos_mt > 15:
                    break
                continue
                
            vazios_consecutivos_mt = 0
            chave_mt = (limpar_texto(mercado_raw), limpar_texto(marca_raw), limpar_texto(serie_raw))
                        
            if chave_mt in dados_extraidos:
                valores_rt = dados_extraidos[chave_mt]['RT']
                for offset in range(12):
                    aba_mt.cell(row=linha, column=16 + offset).value = valores_rt[offset]
                
                linhas_atualizadas_total += 1

    saida_virtual = BytesIO()
    wb_mt.save(saida_virtual)
    saida_virtual.seek(0)
    
    return saida_virtual, linhas_atualizadas_total

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
        help="Transfere os dados do DR para o material de trabalho."
    )
    st.markdown("---")
    
    col1, col2 = st.columns(2)
    with col1:
        arquivo_dr = st.file_uploader("📥 Upload do Arquivo DR", type=["xlsx", "xlsm"])
    with col2:
        arquivo_mt = st.file_uploader("📥 Upload do Material de Trabalho", type=["xlsx", "xlsm"])
        
    st.markdown("### 🎯 Parâmetros da Transferência")
    
    if arquivo_dr is not None and arquivo_mt is not None:
        try:
            abas_dr = pd.ExcelFile(arquivo_dr).sheet_names
            
            col_ano, col_mercado, col_marca = st.columns(3)
            with col_ano:
                anos_selecionados = st.multiselect("Selecione os Anos", options=["2026", "2027"])
            with col_mercado:
                mercados_selecionados = st.multiselect("Mercados", options=["BRA", "ARG", "OSA"])
            with col_marca:
                marcas_selecionadas = st.multiselect("Marcas", options=["FE", "MF", "VT"])

            st.markdown("---")
            st.write("📌 **Regra de Negócio (Estoque)**")
            
            meses_dict = {
                "Janeiro": 1, "Fevereiro":
