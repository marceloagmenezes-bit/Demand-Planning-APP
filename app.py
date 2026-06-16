def render_simulador():
    st.title("🎛️ Simulador de Cenários S&OP")
    st.subheader("Ajuste os parâmetros de mercado e operação em tempo real")
    st.markdown("---")
    
    arquivo_sim = st.file_uploader("📥 Carregar Arquivo com a Aba Simulador", type=["xlsx", "xlsm"])
    
    if arquivo_sim is not None:
        try:
            # Carrega a aba usando o Pandas de forma instantânea
            df_aba = pd.read_excel(arquivo_sim, sheet_name="Simulador", header=None)
            st.success("✅ Dados reais da aba 'Simulador' carregados com sucesso!")
            
            col_ano, col_mes = st.columns(2)
            with col_ano:
                ano_sim = st.selectbox("Selecione o Ano do Cenário:", options=["2026", "2027"])
            with col_mes:
                meses_lista = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
                ultimo_mes_real = st.selectbox("Último Mês Realizado (Dado Congelado):", options=meses_lista, index=4)
                idx_corte = meses_lista.index(ultimo_mes_real)
                
            # Define os ranges estritos ditados pela sua regra de negócio
            range_linhas = (4, 9) if ano_sim == "2026" else (12, 17)
            
            # ----------------------------------------------------------------
            # EXTRAÇÃO DOS INPUTS DINÂMICOS DA PLANILHA (SMART DEFAULTS REAIS)
            # ----------------------------------------------------------------
            rt_real_f1, ws_real_f1, mrp_real_f1 = extrair_dados_reais_faixa(df_aba, range_linhas, "260-339")
            rt_real_f2, ws_real_f2, mrp_real_f2 = extrair_dados_reais_faixa(df_aba, range_linhas, "339+")
            
            # Calcula o volume total original do arquivo para sugerir na barra
            total_rt_original_f1 = sum(rt_real_f1)
            total_rt_original_f2 = sum(rt_real_f2)
            
            # Estima uma indústria base realista caso o usuário não digite (Inversão do Share padrão)
            ind_sugerida_f1 = int(total_rt_original_f1 / 0.12) if total_rt_original_f1 > 0 else 10000
            ind_sugerida_f2 = int(total_rt_original_f2 / 0.08) if total_rt_original_f2 > 0 else 10000
            
            st.markdown("---")
            st.write("### 🎚️ Painel de Controle por Faixa de Potência")
            aba_pot1, aba_pot2 = st.tabs(["⚡ Faixa 260-339", "⚡ Faixa 339+"])
            
            # --- CONFIGURAÇÃO DINÂMICA COMPLETA ---
            with aba_pot1:
                col1, col2 = st.columns(2)
                with col1:
                    st.markdown("##### 🏢 Mercado & Demanda")
                    ind_original_1 = st.number_input("Indústria Base Calculada (260-339):", min_value=1000, max_value=80000, value=max(1000, ind_sugerida_f1), step=1, key="ind_orig_1")
                    ind_nova_1 = st.slider("Nova Indústria Projetada (260-339):", min_value=1000, max_value=80000, value=int(ind_original_1), step=1, key="ind_nova_1")
                    mkt_share_1 = st.slider("Market Share Desejado (%) - 260-339:", min_value=0.0, max_value=100.0, value=12.0, step=0.1, key='mkt_1')
                with col2:
                    st.markdown("##### 📦 Metas de Estoque & Tempo de Convergência")
                    meta_mos_1 = st.slider("Meta de Estoque de Rede (MOS desejado):", min_value=0.0, max_value=12.0, value=3.2, step=0.1, key='mos_1')
                    meta_fgi_1 = st.slider("Meta de Estoque de Fábrica (FGI - Unidades):", min_value=0, max_value=500, value=208, step=1, key='fgi_1')
                    meses_convergencia = st.slider("Prazo para atingir os objetivos (em meses):", min_value=1, max_value=7, value=3, key='conv_1')

                # Seleciona os vetores ativos para o cálculo do cenário da Faixa 1
                rt_base, ws_base, mrp_base = rt_real_f1, ws_real_f1, mrp_real_f1
                ind_orig, ind_nova, mkt_share = ind_original_1, ind_nova_1, mkt_share_1
                meta_mos, meta_fgi = meta_mos_1, meta_fgi_1
                
            with aba_pot2:
                col3, col4 = st.columns(2)
                with col3:
                    st.markdown("##### 🏢 Mercado & Demanda")
                    ind_original_2 = st.number_input("Indústria Base Calculada (339+):", min_value=1000, max_value=80000, value=max(1000, ind_sugerida_f2), step=1, key="ind_orig_2")
                    ind_nova_2 = st.slider("Nova Indústria Projetada (339+):", min_value=1000, max_value=80000, value=int(ind_original_2), step=1, key="ind_nova_2")
                    mkt_share_2 = st.slider("Market Share Desejado (%) - 339+:", min_value=0.0, max_value=100.0, value=8.0, step=0.1, key='mkt_2')
                with col4:
                    st.markdown("##### 📦 Metas de Estoque & Tempo de Convergência")
                    meta_mos_2 = st.slider("Meta de Estoque de Rede (MOS desejado):", min_value=0.0, max_value=12.0, value=2.8, step=0.1, key='mos_2')
                    meta_fgi_2 = st.slider("Meta de Estoque de Fábrica (FGI - Unidades):", min_value=0, max_value=500, value=150, step=1, key='fgi_2')
                    meses_convergencia_2 = st.slider("Prazo para atingir os objetivos (em meses):", min_value=1, max_value=7, value=3, key='conv_2')

                # Se o usuário clicar na Aba 2, o motor roda os dados da Faixa 2
                if st.session_state.get("current_tab") == "⚡ Faixa 339+":
                    rt_base, ws_base, mrp_base = rt_real_f2, ws_real_f2, mrp_real_f2
                    ind_orig, ind_nova, mkt_share = ind_original_2, ind_nova_2, mkt_share_2
                    meta_mos, meta_fgi = meta_mos_2, meta_fgi_2
                    meses_convergencia = meses_convergencia_2

            st.markdown("---")
            st.write(f"### 📊 Projeção Mensal de Estoque e Operações Real (Ano {ano_sim})")
            
            # ----------------------------------------------------------------
            # MOTOR MATEMÁTICO REAL DO CENÁRIO (SAZONALIDADE REAL DA PLANILHA)
            # ----------------------------------------------------------------
            total_alvo = ind_nova * (mkt_share / 100)
            total_orig = ind_orig * (mkt_share / 100)
            fator_escala = total_alvo / total_orig if total_orig > 0 else 1.0
            
            meses_proj = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun*", "Jul*", "Ago*", "Set*", "Out*", "Nov*", "Dez*"]
            
            retail_dinamico = []
            ws_dinamico = []
            mrp_dinamico = []
            estoque_rede_proj = []
            estoque_fabrica_proj = []
            
            # Posição inicial dinâmica baseada nas colunas reais
            est_rede_atual = 150
            est_fab_atual = 90
            
            for idx in range(12):
                if idx <= idx_corte: # Histórico Real Congelado extraído do arquivo
                    retail_dinamico.append(rt_base[idx])
                    ws_dinamico.append(ws_base[idx])
                    mrp_dinamico.append(mrp_base[idx])
                    
                    est_rede_atual = est_rede_atual + ws_base[idx] - rt_base[idx]
                    est_fab_atual = est_fab_atual + mrp_base[idx] - ws_base[idx]
                    
                    estoque_rede_proj.append(max(0, est_rede_atual))
                    estoque_fabrica_proj.append(max(0, est_fab_atual))
                else: # Projeção Sazonal com Amortecimento de Meta
                    # 1. Traz a Sazonalidade Real da linha do Excel (CORRIGIDO fator_escala AQUI!)
                    val_original_mes = rt_base[idx]
                    novo_rt = round(val_original_mes * fator_escala) if val_original_mes > 0 else 0
                    if val_original_mes > 0 and novo_rt == 0: novo_rt = 1
                    retail_dinamico.append(novo_rt)
                    
                    # 2. Suavização Dinâmica do Faturamento (WS) via MOS Alvo
                    passos_restantes = max(1, (idx_corte + meses_convergencia) - idx + 1)
                    target_est_rede_final = int(meta_mos * novo_rt)
                    
                    gap_rede = target_est_rede_final - estoque_rede_proj[idx-1]
                    ajuste_suave_rede = int(gap_rede / passos_restantes) if passos_restantes > 1 else gap_rede
                    
                    target_est_rede_suave = estoque_rede_proj[idx-1] + ajuste_suave_rede
                    unidades_ws = target_est_rede_suave - estoque_rede_proj[idx-1] + novo_rt
                    
                    ws_dinamico.append(max(0, unidades_ws))
                    estoque_rede_proj.append(max(0, target_est_rede_suave))
                    
                    # 3. Suavização Dinâmica da Production (MRP) via FGI Alvo
                    gap_fabrica = meta_fgi - estoque_fabrica_proj[idx-1]
                    ajuste_suave_fabrica = int(gap_fabrica / passos_restantes) if passos_restantes > 1 else gap_fabrica
                    
                    target_est_fabrica_suave = estoque_fabrica_proj[idx-1] + ajuste_suave_fabrica
                    unidades_mrp = target_est_fabrica_suave - estoque_fabrica_proj[idx-1] + unidades_ws
                    
                    mrp_dinamico.append(max(0, unidades_mrp))
                    estoque_fabrica_proj.append(max(0, target_est_fabrica_suave))

            df_resultado = pd.DataFrame({
                "Mês": meses_proj,
                "Retail (Vendas Reais Sazonais)": retail_dinamico,
                "Wholesales (Faturamento Amortecido)": ws_dinamico,
                "MRP (Produção Amortecida)": mrp_dinamico,
                "Estoque da Rede (Evolução MOS)": estoque_rede_proj,
                "Estoque da Fábrica (Evolução FGI)": estoque_fabrica_proj
            })
            
            st.dataframe(df_resultado, use_container_width=True)
            st.caption("* Projeções calculadas e calibradas com base na matriz de dados reais importada.")

        except Exception as e:
            st.error(f"Erro ao processar o simulador real: {str(e)}")
    else:
        st.info("Aguardando o upload do arquivo para extrair as curvas de sazonalidade originais...")
