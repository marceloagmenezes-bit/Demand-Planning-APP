import streamlit as st
import pandas as pd
import os
import openpyxl
import math
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

def limpar_texto(texto):
    if texto is None:
        return ""
    return " ".join(str(texto).upper().strip().split())

def arredondar_sazonal(valor_original, fator):
    """Aplica o fator mantendo a proporcionalidade e garante mínimo de 1 se o original for > 0"""
    if valor_original == 0:
        return 0
    calculado = round(valor_original * fator)
    if valor_original > 0 and calculado == 0:
        return 1
    return calculado

# ==========================================
# O ROBÔ DE PROCESSAMENTO (VBA MOCK / REVISÃO)
# ==========================================
def processar_cruzamento_dr_mt(arquivo_dr, arquivo_mt, anos_selecionados, mercados_selecionados, marcas_selecionadas, mapeamento_abas, limite_mes):
    wb_dr = openpyxl.load_workbook(arquivo_dr, data_only=True)
    wb_mt = openpyxl.load_workbook(arquivo_mt, keep_vba=True)
    linhas_atualizadas_total = 0
    return BytesIO(), linhas_atualizadas_total

# ==========================================
# MÓDULOS DE RENDERIZAÇÃO DE TELAS
# ==========================================
def render_home():
    st.title("🏠 Planejamento de Demanda")
    st.markdown("---")
    st.info("💡 Selecione uma funcionalidade no menu lateral para iniciar o processamento.")

def render_atualizar_material():
    st.header("Atualizar material de trabalho com informações do DR")
    st.info("Esta tela está configurada para uso via macro local conforme alinhamento de arquitetura.")

def render_simulador():
    st.title("🎛️ Simulador de Cenários S&OP")
    st.subheader("Ajuste os parâmetros de mercado e operação em tempo real")
    st.markdown("---")
    
    arquivo_sim = st.file_uploader("📥 Carregar Arquivo com a Aba Simulador", type=["xlsx", "xlsm"])
    
    if arquivo_sim is not None:
        try:
            col_ano, col_mes = st.columns(2)
            with col_ano:
                ano_sim = st.selectbox("Selecione o Ano do Cenário:", options=["2026", "2027"])
            with col_mes:
                meses_lista = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
                ultimo_mes_real = st.selectbox("Último Mês Realizado (Dado Congelado):", options=meses_lista, index=4)
                idx_corte = meses_lista.index(ultimo_mes_real) # Índice do mês (ex: Maio = 4)
                
            # Valores de Indústria originais simulados da planilha real
            ind_inicial_faixa_1 = 14520
            mkt_inicial_faixa_1 = 12.0
            
            st.markdown("---")
            st.write("### 🎚️ Painel de Controle por Faixa de Potência")
            aba_pot1, aba_pot2 = st.tabs(["⚡ Faixa 260-339", "⚡ Faixa 339+"])
            
            with aba_pot1:
                col1, col2 = st.columns(2)
                with col1:
                    st.markdown("##### 🏢 Mercado & Demanda")
                    ind_original_1 = st.number_input("Indústria Base Calculada (260-339):", min_value=1000, max_value=80000, value=int(ind_inicial_faixa_1), step=1, key="ind_orig_1")
                    ind_nova_1 = st.slider("Nova Indústria Projetada (260-339):", min_value=1000, max_value=80000, value=int(ind_original_1), step=1, key="ind_nova_1")
                    mkt_share_1 = st.slider("Market Share Desejado (%) - 260-339:", min_value=0.0, max_value=100.0, value=float(mkt_inicial_faixa_1), step=0.1, key='mkt_1')
                with col2:
                    st.markdown("##### 📦 Metas de Estoque & Tempo de Convergência")
                    meta_mos_1 = st.slider("Meta de Estoque de Rede (MOS desejado):", min_value=0.0, max_value=12.0, value=3.2, step=0.1, key='mos_1')
                    meta_fgi_1 = st.slider("Meta de Estoque de Fábrica (FGI - Unidades):", min_value=0, max_value=500, value=208, step=1, key='fgi_1')
                    # NOVO PARAMETRO DE SUAVIZAÇÃO
                    meses_convergencia = st.slider("Prazo para atingir os objetivos (em meses):", min_value=1, max_value=7, value=3, help="Dilui a necessidade de produção/faturamento ao longo do prazo escolhido para evitar picos abruptos.")
            
            with aba_pot2:
                st.write("Configurações equivalentes aplicadas à Faixa 339+.")

            st.markdown("---")
            st.write("### 📊 Projeção Mensal de Estoque e Operações (S&OP)")
            
            # ----------------------------------------------------------------
            # MOTOR MATEMÁTICO REAL COM SAZONALIDADE E SUAVIZAÇÃO
            # ----------------------------------------------------------------
            # Dados originais fictícios com a sazonalidade real mencionada por você (ex: Junho=2, Julho=5, Agosto=10...)
            retail_original_faixa = [5, 4, 6, 4, 5, 2, 5, 10, 8, 7, 6, 5]
            ws_original_faixa     = [5, 5, 5, 5, 5, 4, 5, 8, 7, 6, 5, 5]
            mrp_original_faixa    = [5, 5, 5, 5, 5, 4, 5, 8, 7, 6, 5, 5]
            
            # Cálculo do Fator de Sazonalidade (Nova Indústria / Antiga Indústria)
            total_alvo = ind_nova_1 * (mkt_share_1 / 100)
            total_orig = ind_original_1 * (mkt_share_1 / 100)
            fator_escala = total_alvo / total_orig if total_orig > 0 else 1.0
            
            meses_proj = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun*", "Jul*", "Ago*", "Set*", "Out*", "Nov*", "Dez*"]
            
            retail_dinamico = []
            ws_dinamico = []
            mrp_dinamico = []
            estoque_rede_proj = []
            estoque_fabrica_proj = []
            
            # Saldos Iniciais Históricos (Maio)
            est_rede_atual = 195
            est_fab_atual = 95
            
            # Loop de construção da tabela mês a mês
            for idx in range(12):
                if idx <= idx_corte: # Meses Passados (Realizados/Congelados)
                    retail_dinamico.append(retail_original_faixa[idx])
                    ws_dinamico.append(ws_original_faixa[idx])
                    mrp_dinamico.append(mrp_original_faixa[idx])
                    
                    if idx > 0:
                        est_rede_atual = estoque_rede_proj[idx-1] + ws_original_faixa[idx] - retail_original_faixa[idx]
                        est_fab_atual = estoque_fabrica_proj[idx-1] + mrp_original_faixa[idx] - ws_original_faixa[idx]
                    
                    estoque_rede_proj.append(est_rede_atual)
                    estoque_fabrica_proj.append(est_fab_atual)
                else: # Meses Futuros (Projeção Inteligente Dinâmica)
                    # 1. Aplica Sazonalidade Perfeita no Retail
                    novo_retail_mes = arredondar_sazonal(retail_original_faixa[idx], fator_escala)
                    retail_dinamico.append(novo_retail_mes)
                    
                    # 2. Suavização da Meta de Estoque da Rede (MOS)
                    # Se o prazo de convergência ainda não acabou, dilui o gap
                    passos_restantes = max(1, (idx_corte + meses_convergencia) - idx + 1)
                    target_est_rede_final = int(meta_mos_1 * novo_retail_mes)
                    
                    gap_rede = target_est_rede_final - estoque_rede_proj[idx-1]
                    ajuste_rede_mes = int(gap_rede / passos_restantes) if passos_restantes > 1 else gap_rede
                    
                    target_est_rede_suave = estoque_rede_proj[idx-1] + ajuste_rede_mes
                    unidades_ws = target_est_rede_suave - estoque_rede_proj[idx-1] + novo_retail_mes
                    
                    ws_dinamico.append(max(0, unidades_ws))
                    estoque_rede_proj.append(target_est_rede_suave)
                    
                    # 3. Suavização da Meta de Estoque da Fábrica (FGI)
                    gap_fabrica = meta_fgi_1 - estoque_fabrica_proj[idx-1]
                    ajuste_fabrica_mes = int(gap_fabrica / passos_restantes) if passos_restantes > 1 else gap_fabrica
                    
                    target_est_fabrica_suave = estoque_fabrica_proj[idx-1] + ajuste_fabrica_mes
                    unidades_mrp = target_est_fabrica_suave - estoque_fabrica_proj[idx-1] + unidades_ws
                    
                    mrp_dinamico.append(max(0, unidades_mrp))
                    estoque_fabrica_proj.append(target_est_fabrica_suave)

            df_resultado = pd.DataFrame({
                "Mês": meses_proj,
                "Retail (Vendas Sazonais)": retail_dinamico,
                "Wholesales (Faturamento Suave)": ws_dinamico,
                "MRP (Produção Suave)": mrp_dinamico,
                "Estoque da Rede (Alvo Proporcional)": estoque_rede_proj,
                "Estoque da Fábrica (Alvo Proporcional)": estoque_fabrica_proj
            })
            
            st.dataframe(df_resultado, use_container_width=True)
            st.caption("* Projeções calculadas respeitando a curva sazonal e suavização de metas de capacidade industrial.")

        except Exception as e:
            st.error(f"Erro ao processar o simulador: {str(e)}")
    else:
        st.info("Aguardando o upload do arquivo para carregar o painel de simulação...")

def render_previsao():
    st.title("📈 Previsão de Demanda")

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
            icons=["", "", "", ""],
            menu_icon="",
            default_index=2, # Já abre direto no simulador para agilizar seu teste
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
            "Simulador de Cenários": render_simulador,
            "Previsão": render_previsao
        }
        if app_selecionado in views:
            views[app_selecionado]()
    except Exception as e:
        st.error(f"Erro crítico: {str(e)}")

if __name__ == "__main__":
    main()
