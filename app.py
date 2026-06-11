import streamlit as st
import pandas as pd
import os
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
    """
    Injeta CSS personalizado para forçar o fundo branco no menu lateral
    e adicionar uma linha de divisão cinza escura, mantendo a capacidade de redimensionamento.
    """
    estilo_css = """
    <style>
        /* Força o fundo branco no menu lateral */
        [data-testid="stSidebar"] {
            background-color: #FFFFFF !important;
            border-right: 2px solid #4F4F4F !important; /* Linha cinza escura */
        }
        
        /* Ajuste fino para garantir que o container interno também fique branco */
        [data-testid="stSidebar"] > div:first-child {
            background-color: #FFFFFF !important;
        }
    </style>
    """
    st.markdown(estilo_css, unsafe_allow_html=True)

# Aplica o estilo logo no início da execução
aplicar_estilo_corporativo()

# ==========================================
# MÓDULOS DE RENDERIZAÇÃO DE TELAS
# ==========================================
def render_home():
    st.title("🏠 Planejamento de Demanda")
    st.subheader("Bem-vindo ao Portal de Demand Planning")
    st.markdown("---")
    st.info("💡 Selecione uma funcionalidade no menu lateral para iniciar o processamento.")

def render_carga_dados():
    st.title("📊 Carga de Dados")
    st.subheader("Upload de Arquivos para Análise de Demanda")
    st.markdown("---")
    st.write("Suba o histórico de vendas do departamento nos formatos **.xlsx (Excel)** ou **.csv**.")
    
    arquivo_carregado = st.file_uploader(
        label="Arraste ou selecione seu arquivo aqui",
        type=["xlsx", "csv"],
        help="Limite de 200MB por arquivo."
    )
    
    if arquivo_carregado is not None:
        try:
            nome_arquivo = arquivo_carregado.name
            with st.spinner(f"Processando arquivo {nome_arquivo}..."):
                if nome_arquivo.endswith('.csv'):
                    df_dados = pd.read_csv(arquivo_carregado)
                else:
                    df_dados = pd.read_excel(arquivo_carregado)
            
            st.success("✅ Arquivo carregado e processado com sucesso!")
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric(label="Total de Linhas", value=f"{df_dados.shape[0]:,}")
            with col2:
                st.metric(label="Total de Colunas", value=df_dados.shape[1])
            with col3:
                memoria_mb = df_dados.memory_usage(deep=True).sum() / (1024 * 1024)
                st.metric(label="Uso de Memória", value=f"{memoria_mb:.2f} MB")
            
            st.markdown("### 🔍 Pré-visualização dos Dados")
            st.dataframe(df_dados.head(10), use_container_width=True)
            
        except Exception as e:
            st.error("❌ Erro ao processar o arquivo. Certifique-se de que o formato está correto.")
            st.warning(f"Detalhe técnico: {str(e)}")

def render_previsao():
    st.title("📈 Previsão de Demanda")
    st.subheader("Modelos Estatísticos e Projeções")
    st.markdown("---")
    st.write("Área destinada ao cálculo de Forecast e projeções futuras.")

# ==========================================
# COMPONENTE DO MENU LATERAL
# ==========================================
def build_sidebar():
    with st.sidebar:
        # 1. LOGO DA EMPRESA/DEPARTAMENTO (Atualizado para .jpg)
        nome_logo = "logo.jpg"
        
        if os.path.exists(nome_logo):
            st.image(nome_logo, use_container_width=True)
        else:
            st.markdown("### 🏢 DEP. DE PLANEJAMENTO")
            
        st.caption("Ambiente Nuvem Protegido")
        st.markdown("---")
        
        # 2. MENU DE NAVEGAÇÃO (Sem Ícones)
        selected_app = option_menu(
            menu_title="Navegação",
            options=["Home", "Carga de Dados", "Previsão"],
            icons=["", "", ""],  # Ícones removidos
            menu_icon="",        # Ícone principal removido
            default_index=0,
            styles={
                "container": {"padding": "5px!", "background-color": "#FFFFFF"}, # Fundo branco no menu
                "nav-link": {
                    "font-size": "15px", 
                    "text-align": "left", 
                    "margin":"0px", 
                    "--hover-color": "#F0F2F6", # Cor de destaque ao passar o mouse
                    "color": "#31333F" # Cor da fonte padrão cinza escuro
                },
                "nav-link-selected": {
                    "background-color": "#0078D4", # Azul corporativo para o item selecionado
                    "color": "white"
                },
            }
        )
    return selected_app

# ==========================================
# FLUXO PRINCIPAL (EXECUÇÃO)
# ==========================================
def main():
    try:
        app_selecionado = build_sidebar()
        
        views = {
            "Home": render_home,
            "Carga de Dados": render_carga_dados,
            "Previsão": render_previsao
        }
        
        if app_selecionado in views:
            views[app_selecionado]()
            
    except Exception as e:
        st.error(f"Ocorreu um erro crítico na aplicação: {str(e)}")

if __name__ == "__main__":
    main()
