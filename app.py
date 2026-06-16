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
    """Remove espaços extras e padroniza para maiúsculo para garantir o cruzamento exato."""
    if texto is None:
        return ""
    return " ".join(str(texto).upper().strip().split())

# ==========================================
# O ROBÔ DE PROCESSAMENTO (PANDAS + OPENPYXL)
# ==========================================
def processar_cruzamento_dr_mt(arquivo_dr, arquivo_mt, anos_selecionados, mercados_selecionados, marcas_selecionadas, mapeamento_abas, limite_mes):
    """
    Versão Turbo e Robusta: Lê os arquivos, cruza a chave e injeta os dados.
    Usa Early Exit (Falso Fim) para ler rápido e limpeza de string para garantir o Match.
    """
    # Carregamento dos arquivos
    wb_dr = openpyxl.load_workbook(arquivo_dr, data_only=True)
    wb_mt = openpyxl.load_workbook(arquivo_mt, keep_vba=True)
    
    linhas_atualizadas_total = 0
    
    # Prepara as listas de filtros limpas para o Match
    mercados_sel_clean = [limpar_texto(m) for m in mercados_selecionados]
    marcas_sel_clean = [limpar_texto(m) for m in marcas_selecionadas]

    for ano in anos_selecionados:
        aba_dr_nome = mapeamento_abas[ano]
        aba_dr = wb_dr[aba_dr_nome]
        
        # Verifica se o MT tem a aba exata do ano
        if ano not in wb_mt.sheetnames:
            raise ValueError(f"A aba '{ano}' não foi encontrada no Material de Trabalho.")
        aba_mt = wb_mt[ano]
        
        # Função Inteligente para achar colunas (Mesmo se for 'SÉRIE(GRUPO)' ou 'MER.')
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
        
        # Validação de Segurança
        if not all(k in cols_dr for k in ['MERCADO', 'MARCA', 'SERIE']):
            raise ValueError(f"Cabeçalhos (Mercado, Marca, Série) não encontrados na linha 4 do arquivo DR ({aba_dr_nome}).")
            
        dados_extraidos = {}
        
        # ==========================================
        # LEITURA RÁPIDA DO DR (COM EARLY EXIT)
        # ==========================================
        vazios_consecutivos = 0
        for linha in range(5, aba_dr.max_row + 1):
            mercado_raw = aba_dr.cell(row=linha, column=cols_dr['MERCADO']).value
            marca_raw = aba_dr.cell(row=linha, column=cols_dr['MARCA']).value
            serie_raw = aba_dr.cell(row=linha, column=cols_dr['SERIE']).value
            produto_raw = aba_dr.cell(row=linha, column=cols_dr.get('PRODUTO', 0)).value
            
            # Trava de Performance (Se achar 15 linhas vazias seguidas, a tabela acabou)
            if not serie_raw:
                vazios_consecutivos += 1
                if vazios_consecutivos > 15:
                    break
                continue
            
            vazios_consecutivos = 0 # Reseta o contador se achou linha preenchida
            
            # Limpeza
            mercado = limpar_texto(mercado_raw)
            marca = limpar_texto(marca_raw)
            serie = limpar_texto(serie_raw)
            produto = limpar_texto(produto_raw)
            
            if "TOTAL" in produto:
                continue
                
            if mercado in mercados_sel_clean and marca in marcas_sel_clean:
                chave = (mercado, marca, serie)
                
                # Coleta RT (12 meses totais)
                valores_rt = []
                for offset in range(12):
                    valor_celula = aba_dr.cell(row=linha, column=16 + offset).value
                    valores_rt.append(valor_celula if valor_celula is not None else 0)
                
                dados_extraidos[chave] = {
                    'RT': valores_rt
                }

        # ==========================================
        # INJEÇÃO RÁPIDA NO MT (COM EARLY EXIT)
        # ==========================================
        vazios_consecutivos_mt = 0
        for linha in range(5, aba_mt.max_row + 1):
            # Fallback seguro caso ele não ache a coluna no MT (D, E, G conforme imagem)
            col_mer = cols_mt.get('MERCADO', 4)
            col_mar = cols_mt.get('MARCA', 5)
            col_ser = cols_mt.get('SERIE', 7) # G é a 7ª coluna
            
            mercado_raw = aba_mt.cell(row=linha, column=col_mer).value
            marca_raw = aba_mt.cell(row=linha, column=col_mar).value
            serie_raw = aba_mt.cell(row=linha, column=col_ser).value
            
            # Trava de Performance
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

    # Prepara o arquivo final
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
                "Janeiro": 1, "Fevereiro": 2, "Março": 3, "Abril": 4, 
                "Maio": 5, "Junho": 6, "Julho": 7, "Agosto": 8, 
                "Setembro": 9, "Outubro": 10, "Novembro": 11, "Dezembro": 12
            }
            
            mes_selecionado = st.selectbox(
                "Selecione o último mês realizado (Real):", 
                options=list(meses_dict.keys()), 
                index=4
            )
            limite_mes = meses_dict[mes_selecionado]

            st.markdown("---")
            
            mapeamento_abas = {}
            if anos_selecionados:
                col_abas = st.columns(len(anos_selecionados))
                for idx, ano in enumerate(anos_selecionados):
                    with col_abas[idx]:
                        mapeamento_abas[ano] = st.selectbox(
                            f"Aba do DR com os dados de {ano}:", 
                            options=abas_dr,
                            key=f"aba_{ano}"
                        )

            st.markdown("<br>", unsafe_allow_html=True)
            
            if st.button("🚀 Executar Transferência de Dados", type="primary", use_container_width=True):
                if not anos_selecionados or not mercados_selecionados or not marcas_selecionadas:
                    st.warning("⚠️ Selecione ao menos um Ano, um Mercado e uma Marca.")
                else:
                    try:
                        with st.spinner("Modo Turbo ativado. Processando cruzamento de dados..."):
                            arquivo_pronto, qtd_atualizadas = processar_cruzamento_dr_mt(
                                arquivo_dr=arquivo_dr,
                                arquivo_mt=arquivo_mt,
                                anos_selecionados=anos_selecionados,
                                mercados_selecionados=mercados_selecionados,
                                marcas_selecionadas=marcas_selecionadas,
                                mapeamento_abas=mapeamento_abas,
                                limite_mes=limite_mes
                            )
                        
                        if qtd_atualizadas > 0:
                            st.success(f"✅ Sucesso! **{qtd_atualizadas} séries** foram cruzadas e atualizadas no Material de Trabalho.")
                            
                            extensao_mt = os.path.splitext(arquivo_mt.name)[1].lower()
                            mime_type = "application/vnd.ms-excel.sheet.macroEnabled.12" if extensao_mt == ".xlsm" else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            
                            st.download_button(
                                label=f"⬇️ Baixar Material de Trabalho Atualizado ({extensao_mt})",
                                data=arquivo_pronto,
                                file_name=f"Material_de_Trabalho_Atualizado{extensao_mt}",
                                mime=mime_type,
                                use_container_width=True
                            )
                        else:
                            st.warning("⚠️ Processo finalizado, mas **NENHUMA** série foi atualizada. Verifique se o DR realmente possui as marcas/mercados que você selecionou e se os nomes das Séries batem exatamente com o MT.")

                    except Exception as e:
                        st.error("❌ Ocorreu um erro durante o cruzamento das planilhas.")
                        st.write(f"Detalhe técnico: {str(e)}")
                        
        except Exception as e:
            st.error("Erro ao ler os arquivos.")
            st.warning(str(e))
    else:
        st.info("Aguardando o upload dos dois arquivos para liberar os seletores...")

def render_previsao():
    st.title("📈 Previsão de Demanda")

# ==========================================
# COMPONENTE DO MENU LATERAL E FLUXO PRINCIPAL
# ==========================================
def build_sidebar():
    with st.sidebar:
        nome_logo = "logo.jpg"
        if os.path.exists(nome_logo):
            st.image(nome_logo, use_container_width=True)
        else:
            st.markdown("### 🏢 DEP. DE PLANEJAMENTO")
            
        st.caption("Ambiente Nuvem Protegido")
        st.markdown("---")
        
        selected_app = option_menu(
            menu_title="Navegação",
           options=["Home", "Atualizar Material DR", "Simulador de Cenários", "Previsão"],
            icons=["", "", ""],
            menu_icon="",
            default_index=0,
            styles={
                "container": {"padding": "5px!", "background-color": "#FFFFFF"},
                "nav-link": {"font-size": "15px", "text-align": "left", "margin":"0px", "--hover-color": "#F0F2F6", "color": "#31333F"},
                "nav-link-selected": {"background-color": "#0078D4", "color": "white"},
            }
        )
    return selected_app

def main():
    try:
        app_selecionado = build_sidebar()
        views = {
    "Home": render_home,
    "Atualizar Material DR": render_atualizar_material,
    "Simulador de Cenários": render_simulador, # <- Conecta a nova tela
    "Previsão": render_previsao
}
        if app_selecionado in views:
            views[app_selecionado]()
    except Exception as e:
        st.error(f"Erro crítico: {str(e)}")

if __name__ == "__main__":
    main()
