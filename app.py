import streamlit as st
import os
from streamlit_option_menu import option_menu

# Configuração da página
st.set_page_config(
    page_title="Demand Planning APP",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Telas do App (Módulos)
def render_home():
    st.title("🏠 Planejamento de Demanda")
    st.subheader("Bem-vindo ao Portal de Demand Planning")
    st.markdown("---")
    st.info("Selecione uma funcionalidade no menu lateral para iniciar.")

def render_carga_dados():
    st.title("📊 Carga de Dados")
    st.subheader("Upload de Arquivos para Análise")
    st.write("Em breve: Área para você arrastar suas planilhas Excel aqui.")

def render_previsao():
    st.title("📈 Previsão de Demanda")
    st.subheader("Modelos e Projeções")
    st.write("Em breve: Gráficos de tendência e previsões futuras.")

# Menu Lateral (25% da tela)
def build_sidebar():
    with st.sidebar:
        st.markdown("### 🏢 DEP. DE PLANEJAMENTO")
        st.caption("Ambiente Nuvem Protegido")
        st.markdown("---")
        
        selected_app = option_menu(
            menu_title="Navegação",
            options=["Home", "Carga de Dados", "Previsão"],
            icons=["house", "cloud-upload", "graph-up-arrow"],
            menu_icon="cast",
            default_index=0,
            styles={
                "container": {"padding": "5px!", "background-color": "#fafafa"},
                "icon": {"color": "#0078D4", "font-size": "16px"}, 
                "nav-link": {"font-size": "14px", "text-align": "left", "margin":"0px", "--hover-color": "#eee"},
                "nav-link-selected": {"background-color": "#0078D4", "color": "white"},
            }
        )
    return selected_app

# Fluxo Principal
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
        st.error(f"Erro na aplicação: {str(e)}")

if __name__ == "__main__":
    main()
