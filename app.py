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

def render_atualizar_material():
    # O parâmetro 'help' cria automaticamente o ícone de '?' com o texto ao passar o mouse!
    st.header(
        "Atualizar material de trabalho com informações do DR", 
        help="Geralmente utilizado no início do ciclo, após atualização dos realizados do mês anterior. Transfere os dados do DR para o material de trabalho: Ao carregar os dois arquivos, selecione as marcas e os mercados desejados, e as informações de RT, WS, MRP e estoques serão transferidas do DR para o Material de trabalho!"
    )
    st.markdown("---")
    
    st.write("Faça o upload dos arquivos base para iniciar o cruzamento de dados.")
    
    # Divide a tela em duas colunas para os uploads ficarem lado a lado
    col1, col2 = st.columns(2)
    with col1:
        arquivo_dr = st.file_uploader("📥 Upload do Arquivo DR", type=["xlsx", "csv"])
    with col2:
        arquivo_mt = st.file_uploader("📥 Upload do Material de Trabalho", type=["xlsx", "csv"])
        
    st.markdown("### 🎯 Filtros de Transferência")
    
    # Divide a tela para os seletores (Iniciam desativados até os arquivos serem carregados)
    col3, col4 = st.columns(2)
    with col3:
        # Nota: Por enquanto os valores são fictícios. Vamos populá-los com o Pandas depois.
        st.multiselect(
            "Selecione as Marcas", 
            options=["Fendt", "Massey Ferguson", "Valtra"], 
            disabled=(arquivo_dr is None), # Só habilita se o arquivo for carregado
            help="Carregue o arquivo DR para liberar as marcas."
        )
    with col4:
        st.multiselect(
            "Selecione os Mercados", 
            options=["Brasil", "Argentina", "Exportação"], 
            disabled=(arquivo_dr is None),
            help="Carregue o arquivo DR para liberar os mercados."
        )
        
    st.markdown("<br>", unsafe_allow_html=True) # Dá um pequeno espaço
    st.button("🚀 Executar Transferência de Dados", type="primary", use_container_width=True)

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
            options=["Home", "Atualizar Material DR", "Previsão"],
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
        
        # Mapeamento de Navegação (Dicionário de Telas)
        views = {
            "Home": render_home,
            "Atualizar Material DR": render_atualizar_material,
            "Previsão": render_previsao
        }
        
        # Renderização dinâmica na área dos 75% da tela
        if app_selecionado in views:
            views[app_selecionado]()
            
    # Este é o bloco que estava faltando ou desalinhado!
    except Exception as e:
        st.error(f"Ocorreu um erro crítico na aplicação: {str(e)}")

if __name__ == "__main__":
    main()
