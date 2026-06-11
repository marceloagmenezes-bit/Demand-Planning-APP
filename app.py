import streamlit as st
import pandas as pd
import os
from streamlit_option_menu import option_menu
import openpyxl
from io import BytesIO

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
    
def processar_cruzamento_dr_mt(arquivo_dr, arquivo_mt, anos_selecionados, mercados_selecionados, marcas_selecionadas, mapeamento_abas):
    """
    Lê os arquivos, cruza a chave (Mercado, Marca, Série) ignorando linhas 'TOTAL',
    e injeta os dados no Material de Trabalho, retornando o arquivo pronto para download.
    """
    # Carrega o DR pegando apenas os valores (ignorando fórmulas)
    wb_dr = openpyxl.load_workbook(arquivo_dr, data_only=True)
    # Carrega o MT preservando toda a formatação e fórmulas originais
    wb_mt = openpyxl.load_workbook(arquivo_mt)
    
    for ano in anos_selecionados:
        aba_dr_nome = mapeamento_abas[ano]
        aba_dr = wb_dr[aba_dr_nome]
        aba_mt = wb_mt[ano] # O MT sempre tem a aba com o nome do ano (2026, 2027)
        
        # 1. FUNÇÃO INTERNA PARA ACHAR AS COLUNAS DINAMICAMENTE NA LINHA 4
        def achar_colunas(aba):
            mapa_cols = {}
            for col in range(1, 50): # Escaneia até a coluna 50
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
        
        # Validação de segurança se achou os cabeçalhos
        if not all(k in cols_dr for k in ['MERCADO', 'MARCA', 'SERIE']):
            raise ValueError(f"Cabeçalhos não encontrados na linha 4 do arquivo DR ({aba_dr_nome}).")
            
        # 2. MAPEAMENTO DOS DADOS DO DR (Criação do Dicionário/Memória)
        dados_extraidos = {}
        
        for linha in range(5, aba_dr.max_row + 1):
            mercado = aba_dr.cell(row=linha, column=cols_dr['MERCADO']).value
            marca = aba_dr.cell(row=linha, column=cols_dr['MARCA']).value
            serie = aba_dr.cell(row=linha, column=cols_dr['SERIE']).value
            produto = aba_dr.cell(row=linha, column=cols_dr.get('PRODUTO', 0)).value
            
            # Limpeza e verificação (Ignora linhas vazias e as que contêm "TOTAL")
            if not serie or not mercado or not marca:
                continue
            if isinstance(produto, str) and "TOTAL" in produto.upper():
                continue
                
            # Verifica se está dentro dos filtros do usuário
            if mercado in mercados_selecionados and marca in marcas_selecionadas:
                chave = (str(mercado).strip(), str(marca).strip(), str(serie).strip())
                
                # Coleta os 12 meses de RT (Assumindo que no DR começa na coluna P=16)
                valores_rt = []
                for offset in range(12):
                    # P(16) a AA(27)
                    valor_celula = aba_dr.cell(row=linha, column=16 + offset).value
                    valores_rt.append(valor_celula if valor_celula is not None else 0)
                
                dados_extraidos[chave] = valores_rt

        # 3. INJEÇÃO DOS DADOS NO MATERIAL DE TRABALHO
        for linha in range(5, aba_mt.max_row + 1):
            mercado = aba_mt.cell(row=linha, column=cols_mt.get('MERCADO', 4)).value # Fallback D(4)
            marca = aba_mt.cell(row=linha, column=cols_mt.get('MARCA', 5)).value   # Fallback E(5)
            serie = aba_mt.cell(row=linha, column=cols_mt.get('SERIE', 6)).value
            
            chave_mt = (str(mercado).strip() if mercado else "", 
                        str(marca).strip() if marca else "", 
                        str(serie).strip() if serie else "")
                        
            if chave_mt in dados_extraidos:
                valores_rt = dados_extraidos[chave_mt]
                
                # Injeta os 12 meses de RT (Assumindo que no MT também começa na P=16)
                # Você pode ajustar esse offset de '16' conforme a real posição do bloco RT no MT
                for offset in range(12):
                    aba_mt.cell(row=linha, column=16 + offset).value = valores_rt[offset]

    # 4. SALVA O ARQUIVO NO FORMATO VIRTUAL PARA DOWNLOAD NA NUVEM
    saida_virtual = BytesIO()
    wb_mt.save(saida_virtual)
    saida_virtual.seek(0)
    
    return saida_virtual
    
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
    
    # Só exibe os filtros se o DR for carregado
    if arquivo_dr is not None and arquivo_mt is not None:
        try:
            # Lê o nome das abas do DR dinamicamente para o usuário escolher
            abas_dr = pd.ExcelFile(arquivo_dr).sheet_names
            
            # Filtros Principais
            col_ano, col_mercado, col_marca = st.columns(3)
            with col_ano:
                anos_selecionados = st.multiselect("Selecione os Anos", options=["2026", "2027"])
            with col_mercado:
                mercados_selecionados = st.multiselect("Mercados", options=["BRA", "ARG", "OSA"])
            with col_marca:
                marcas_selecionadas = st.multiselect("Marcas", options=["FE", "MF", "VT"])

            st.markdown("---")
            
            # Cria menus dinâmicos para mapear as abas dependendo do ano escolhido
            mapeamento_abas = {}
            if anos_selecionados:
                st.write("📌 **Mapeamento de Abas do DR**")
                col_abas = st.columns(len(anos_selecionados))
                
                for idx, ano in enumerate(anos_selecionados):
                    with col_abas[idx]:
                        aba_escolhida = st.selectbox(
                            f"Qual aba do DR tem os dados de {ano}?", 
                            options=abas_dr,
                            key=f"aba_{ano}"
                        )
                        mapeamento_abas[ano] = aba_escolhida

            st.markdown("<br>", unsafe_allow_html=True)
            
       st.markdown("<br>", unsafe_allow_html=True)
            
            # Botão de Execução
            if st.button("🚀 Executar Transferência de Dados", type="primary", use_container_width=True):
                if not anos_selecionados or not mercados_selecionados or not marcas_selecionadas:
                    st.warning("⚠️ Por favor, selecione ao menos um Ano, um Mercado e uma Marca antes de executar.")
                else:
                    try:
                        with st.spinner("Processando cruzamento de dados... Isso pode levar alguns segundos."):
                            # Chama o nosso robô do OpenPyXL
                            arquivo_pronto = processar_cruzamento_dr_mt(
                                arquivo_dr=arquivo_dr,
                                arquivo_mt=arquivo_mt,
                                anos_selecionados=anos_selecionados,
                                mercados_selecionados=mercados_selecionados,
                                marcas_selecionadas=marcas_selecionadas,
                                mapeamento_abas=mapeamento_abas
                            )
                            
                        st.success("✅ Cópia finalizada com sucesso! O layout e as fórmulas foram mantidos.")
                        
                        # Gera o botão para baixar o resultado final
                        st.download_button(
                            label="⬇️ Baixar Material de Trabalho Atualizado",
                            data=arquivo_pronto,
                            file_name="Material_de_Trabalho_Atualizado.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    except Exception as e:
                        st.error("❌ Ocorreu um erro durante o cruzamento das planilhas.")
                        st.write(f"Detalhe técnico: {str(e)}")
                    
        except Exception as e:
            st.error("Erro ao ler os arquivos. Verifique se não estão corrompidos ou protegidos com senha.")
            st.warning(str(e))
    else:
        st.info("Aguardando o upload dos dois arquivos para liberar os seletores...")
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
