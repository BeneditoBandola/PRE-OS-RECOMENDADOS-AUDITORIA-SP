import os
json = __import__('json')
unicodedata = __import__('unicodedata')
pandas = __import__('pandas')
streamlit = st = __import__('streamlit')

# --- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="Auditoria e Ressarcimento - São Paulo",
    layout="wide",
    page_icon="📊",
)

# --- 2. DICIONÁRIO DE LOJAS REAIS (FILIAL RIO PRETO) ---
LOJAS_POR_FILIAL = {
    "SÃO JOSÉ DO RIO PRETO": [
        {"codigo": 9612, "nome": "ANDERSON TOCHETIN RACOES - ME", "cnpj": "13.994.857/0001-16", "cidade": "ARIRANHA", "bairro": "CHACARA SETE ELOS"},
        {"codigo": 6860, "nome": "ADEMIR CANOSSA-CATANDUVA", "cnpj": "03.186.174/0001-89", "cidade": "CATANDUVA", "bairro": "CENTRO"},
        {"codigo": 6859, "nome": "CASA DO CHACAREIRO COM. DE AVES E RACOES LTDA-EPP", "cnpj": "54.871.983/0001-60", "cidade": "CATANDUVA", "bairro": "CENTRO"},
        {"codigo": 4878, "nome": "FAFER COM PROD P/ ANIMAIS E AVES LTDA", "cnpj": "01.781.155/0001-75", "cidade": "MIRASSOL", "bairro": "CENTRO"},
        {"codigo": 11242, "nome": "FRS FERREIRA RACOES", "cnpj": "20.809.189/0001-75", "cidade": "MIRASSOL", "bairro": "JARDIM ALVORADA"},
        {"codigo": 27832, "nome": "JHONATAN CARVALHO TOLEDO LTDA", "cnpj": "63.447.904/0001-80", "cidade": "SANTA ADELIA", "bairro": "CENTRO"},
        {"codigo": 6263, "nome": "ADRIANO JOSE CETRONE S. J. DO RIO PRETO", "cnpj": "05.468.661/0001-60", "cidade": "SAO JOSE DO RIO PRETO", "bairro": "PQ. INDUSTRIAL"},
        {"codigo": 8760, "nome": "DORIS BRUDER PET SHOP LTDA", "cnpj": "62.399.936/0001-94", "cidade": "SAO JOSE DO RIO PRETO", "bairro": "VILA ERCILIA"},
        {"codigo": 12871, "nome": "GBL STORE LTDA", "cnpj": "27.519.711/0001-23", "cidade": "SAO JOSE DO RIO PRETO", "bairro": "PARQUE CELESTE"},
        {"codigo": 13224, "nome": "M.N. STOK DOG PET SHOP LTDA", "cnpj": "26.591.656/0001-10", "cidade": "SAO JOSE DO RIO PRETO", "bairro": "JARDIM RESIDENCIAL VETORASSO"},
        {"codigo": 4964, "nome": "O MUNDO ANIMAL COMERCIO LTDA", "cnpj": "67.731.208/0001-51", "cidade": "SAO JOSE DO RIO PRETO", "bairro": "VL MACENO"}
    ],
    "RIBEIRÃO PRETO": [
        {"codigo": 9001, "nome": "EXEMPLO LOJA RIBEIRÃO PRETO", "cnpj": "00.000.000/0001-00", "cidade": "RIBEIRAO PRETO", "bairro": "CENTRO"}
    ]
}

# --- 3. ESTILO VISUAL E TEMA ---
st.markdown(
    """
<style>
.stApp { background-color: #F8FAFC; color: #1E293B; }
.cabecalho-principal {
    background-color: #FFFFFF;
    border-bottom: 3px solid #C0392B;
    padding: 15px;
    border-radius: 10px;
    margin-bottom: 20px;
    text-align: center;
    box-shadow: 0 2px 5px rgba(0,0,0,0.05);
}
</style>
""",
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="cabecalho-principal"><h2 style="color: #0F172A; margin:0;">📊 Sistema de Auditoria e Ressarcimento (Estado de SP)</h2><p style="color: #64748B; font-size: 13px; margin:5px 0 0 0;">Busca rápida por EAN/Código, controle de preços e salvamento automático em JSON.</p></div>',
    unsafe_allow_html=True,
)

# --- 4. FUNÇÕES DE SUPORTE E CARREGAMENTO DE DADOS ---
@st.cache_data
def carregar_dados():
    arquivo_base = "boneco_com_valor_h.xlsx"
    if not os.path.exists(arquivo_base):
        return None
    try:
        df = pandas.read_excel(arquivo_base, sheet_name="Detalhado")
    except Exception:
        try:
            df = pandas.read_excel(arquivo_base)
        except Exception as e:
            st.error(f"Erro ao abrir a planilha: {e}")
            return None

    df.columns = [str(c).strip().upper() for c in df.columns]
    
    for col in ["CODIGO", "COD. EAN", "SKU"]:
        if col in df.columns:
            df[f"{col}_LIMPO"] = df[col].astype(str).str.replace(".0", "", regex=False).str.strip()
        else:
            df[f"{col}_LIMPO"] = ""
            
    return df

df_produtos = carregar_dados()

if df_produtos is None:
    st.error("⚠️ O arquivo `boneco_com_valor_h.xlsx` não foi encontrado na raiz do projeto.")
    st.stop()

# --- 5. SELEÇÃO DE FILIAL E LOJA (BARRA LATERAL) ---
st.sidebar.header("📍 Configuração da Auditoria")
filial_selecionada = st.sidebar.selectbox("Selecione a Filial:", ["SÃO JOSÉ DO RIO PRETO", "RIBEIRÃO PRETO"])

lista_lojas = LOJAS_POR_FILIAL.get(filial_selecionada, [])
loja_nomes = [f"{l['codigo']} - {l['nome']} ({l['cidade']})" for l in lista_lojas]

loja_escolhida_str = st.sidebar.selectbox("Selecione a Loja:", loja_nomes)

loja_obj = None
if loja_escolhida_str:
    codigo_loja_sel = int(loja_escolhida_str.split(" - ")[0])
    loja_obj = next((l for l in lista_lojas if l['codigo'] == codigo_loja_sel), None)

if loja_obj:
    st.sidebar.markdown("---")
    st.sidebar.markdown(f"**Loja:** {loja_obj['nome']}")
    st.sidebar.markdown(f"**CNPJ:** {loja_obj['cnpj']}")
    st.sidebar.markdown(f"**Cidade/Bairro:** {loja_obj['cidade']} / {loja_obj['bairro']}")

caminho_json = f"auditoria_loja_{loja_obj['codigo']}.json" if loja_obj else "auditoria_loja_padrao.json"

if "auditoria_dados" not in st.session_state:
    st.session_state["auditoria_dados"] = {}

chave_sessao_loja = f"loja_{loja_obj['codigo']}" if loja_obj else "loja_padrao"

if chave_sessao_loja not in st.session_state["auditoria_dados"]:
    df_carregado = None
    if os.path.exists(caminho_json):
        try:
            df_carregado = pandas.read_json(caminho_json, orient="records")
            # Garante conversão de colunas chave para string para evitar erros de leitura
            for c in ['CODIGO_LIMPO', 'COD. EAN_LIMPO', 'SKU_LIMPO']:
                if c in df_carregado.columns:
                    df_carregado[c] = df_carregado[c].astype(str).str.replace(".0", "", regex=False).str.strip()
        except Exception:
            df_carregado = None

    if df_carregado is not None and not df_carregado.empty:
        st.session_state["auditoria_dados"][chave_sessao_loja] = df_carregado
    else:
        # Cria a partir da base garantindo que todas as colunas essenciais existam
        colunas_necessarias = [c for c in ['CODIGO_LIMPO', 'COD. EAN_LIMPO', 'SKU_LIMPO', 'NOME COMERCIAL', 'FAMILIA', 'VALOR_RECOMENDADO_SP'] if c in df_produtos.columns]
        df_init = df_produtos[colunas_necessarias].copy()
        df_init['PREÇO PRATICADO'] = 0.0
        df_init['QUANTIDADE'] = 0
        df_init['STATUS_AUDITORIA'] = "Pendente"
        st.session_state["auditoria_dados"][chave_sessao_loja] = df_init

df_loja_atual = st.session_state["auditoria_dados"][chave_sessao_loja]

# --- 6. ESCOLHA DO MODO DE OPERAÇÃO ---
st.markdown("### 🛠️ Modo de Operação")
modo_operacao = st.radio(
    "Escolha o modo desejado:",
    [
        "1. Cadastrar / Atualizar Preços e Quantidades (Busca por EAN)",
        "2. Relatório de Distorção de Preços (Apenas Preços)",
        "3. Relatório de Cálculo de Ressarcimento (Com Quantidades)"
    ],
    horizontal=True
)

st.markdown("---")

# --- MODO 1: CADASTRO COM BUSCA RÁPIDA POR EAN / CÓDIGO ---
if modo_operacao.startswith("1"):
    st.markdown(f"#### 📝 Cadastro Rápido - {loja_obj['nome'] if loja_obj else ''}")
    st.markdown("💡 **Dica de Campo:** Digite os **últimos 5 números do EAN**, código ou SKU para localizar o produto instantaneamente.")

    filtro_busca = st.text_input("🔍 Pesquisa rápida (Digite parte do EAN, Código ou Nome):", placeholder="Ex: 4731, gastro, kitten...")
    
    df_filtrado = df_loja_atual.copy()
    if filtro_busca:
        termo = filtro_busca.strip().lower()
        
        # Cria condições de busca segura verificando se as colunas existem
        condicoes = []
        if 'COD. EAN_LIMPO' in df_filtrado.columns:
            condicoes.append(df_filtrado['COD. EAN_LIMPO'].str.endswith(termo) | df_filtrado['COD. EAN_LIMPO'].str.contains(termo, na=False))
        if 'CODIGO_LIMPO' in df_filtrado.columns:
            condicoes.append(df_filtrado['CODIGO_LIMPO'].str.endswith(termo) | df_filtrado['CODIGO_LIMPO'].str.contains(termo, na=False))
        if 'SKU_LIMPO' in df_filtrado.columns:
            condicoes.append(df_filtrado['SKU_LIMPO'].str.endswith(termo) | df_filtrado['SKU_LIMPO'].str.contains(termo, na=False))
        if 'NOME COMERCIAL' in df_filtrado.columns:
            condicoes.append(df_filtrado['NOME COMERCIAL'].str.lower().str.contains(termo, na=False))

        if condicoes:
            filtro_final = condicoes[0]
            for c in condicoes[1:]:
                filtro_final = filtro_final | c
            df_filtrado = df_filtrado[filtro_final]

    # Monta dinamicamente as colunas de exibição apenas se existirem no DataFrame
    colunas_possiveis = ['CODIGO_LIMPO', 'COD. EAN_LIMPO', 'NOME COMERCIAL', 'FAMILIA', 'VALOR_RECOMENDADO_SP', 'PREÇO PRATICADO', 'QUANTIDADE']
    colunas_edicao = [c for c in colunas_possiveis if c in df_filtrado.columns]
    
    editado_df = st.data_editor(
        df_filtrado[colunas_edicao],
        column_config={
            "CODIGO_LIMPO": st.column_config.TextColumn("Código", disabled=True),
            "COD. EAN_LIMPO": st.column_config.TextColumn("Cód. EAN", disabled=True),
            "NOME COMERCIAL": st.column_config.TextColumn("Produto", disabled=True),
            "FAMILIA": st.column_config.TextColumn("Família", disabled=True),
            "VALOR_RECOMENDADO_SP": st.column_config.NumberColumn("Preço Rec. (SP)", format="R$ %.2f", disabled=True),
            "PREÇO PRATICADO": st.column_config.NumberColumn("Preço Praticado (R$)", format="R$ %.2f", min_value=0.0, step=0.1),
            "QUANTIDADE": st.column_config.NumberColumn("Qtd (Estoque/Venda)", min_value=0, step=1)
        },
        hide_index=True,
        use_container_width=True,
        key=f"editor_{chave_sessao_loja}"
    )

    if st.button("💾 Salvar e Gravar em JSON", type="primary"):
        for idx, row in editado_df.iterrows():
            codigo_prod = row['CODIGO_LIMPO']
            mask = st.session_state["auditoria_dados"][chave_sessao_loja]['CODIGO_LIMPO'] == codigo_prod
            st.session_state["auditoria_dados"][chave_sessao_loja].loc[mask, 'PREÇO PRATICADO'] = row['PREÇO PRATICADO']
            st.session_state["auditoria_dados"][chave_sessao_loja].loc[mask, 'QUANTIDADE'] = row['QUANTIDADE']
            if row['PREÇO PRATICADO'] > 0:
                st.session_state["auditoria_dados"][chave_sessao_loja].loc[mask, 'STATUS_AUDITORIA'] = "Auditado"
        
        df_para_salvar = st.session_state["auditoria_dados"][chave_sessao_loja]
        df_para_salvar.to_json(caminho_json, orient="records", force_ascii=False, indent=4)
        
        st.success(f"✅ Dados salvos com sucesso e gravados no arquivo `{caminho_json}`!")

# --- MODO 2: RELATÓRIO DE DISTORÇÃO DE PREÇOS ---
elif modo_operacao.startswith("2"):
    st.markdown(f"#### 🔍 Relatório de Distorção de Preços (Apenas Preços)")
    df_atual = st.session_state["auditoria_dados"][chave_sessao_loja]
    df_auditado = df_atual[df_atual['PREÇO PRATICADO'] > 0].copy()

    if df_auditado.empty:
        st.warning("⚠️ Nenhum produto foi auditado com preço praticado ainda.")
    else:
        df_auditado['DIFERENÇA_UNITARIA'] = df_auditado['PREÇO PRATICADO'] - df_auditado['VALOR_RECOMENDADO_SP']
        df_auditado['TIPO_DISTORCAO'] = df_auditado['DIFERENÇA_UNITARIA'].apply(lambda x: "Acima do Recomendado" if x > 0 else ("Abaixo do Recomendado" if x < 0 else "Preço Exato"))

        m1, m2 = st.columns(2)
        m1.metric("SKUs Auditados", len(df_auditado))
        m2.metric("SKUs com Diferença", len(df_auditado[df_auditado['DIFERENÇA_UNITARIA'] != 0]))

        st.markdown("---")
        colunas_rel = [c for c in ['CODIGO_LIMPO', 'COD. EAN_LIMPO', 'NOME COMERCIAL', 'FAMILIA', 'VALOR_RECOMENDADO_SP', 'PREÇO PRATICADO', 'DIFERENÇA_UNITARIA', 'TIPO_DISTORCAO'] if c in df_auditado.columns]
        st.dataframe(
            df_auditado[colunas_rel],
            column_config={
                "CODIGO_LIMPO": "Código",
                "COD. EAN_LIMPO": "EAN",
                "NOME COMERCIAL": "Produto",
                "FAMILIA": "Família",
                "VALOR_RECOMENDADO_SP": st.column_config.NumberColumn("Rec. SP", format="R$ %.2f"),
                "PREÇO PRATICADO": st.column_config.NumberColumn("Praticado", format="R$ %.2f"),
                "DIFERENÇA_UNITARIA": st.column_config.NumberColumn("Diferença Unit.", format="R$ %.2f"),
                "TIPO_DISTORCAO": "Status Distorção"
            },
            hide_index=True,
            use_container_width=True
        )

# --- MODO 3: RELATÓRIO DE CÁLCULO DE RESSARCIMENTO ---
else:
    st.markdown(f"#### 💰 Relatório de Cálculo de Ressarcimento (Com Quantidades)")
    df_atual = st.session_state["auditoria_dados"][chave_sessao_loja]
    df_ressarcimento = df_atual[(df_atual['PREÇO PRATICADO'] > 0) & (df_atual['QUANTIDADE'] > 0)].copy()

    if df_ressarcimento.empty:
        st.warning("⚠️ Preencha os preços e as quantidades no Modo 1 para gerar o ressarcimento.")
    else:
        df_ressarcimento['DIFERENÇA_UNITARIA'] = df_ressarcimento['PREÇO PRATICADO'] - df_ressarcimento['VALOR_RECOMENDADO_SP']
        df_ressarcimento['VALOR_TOTAL_RESSARCIMENTO'] = df_ressarcimento['DIFERENÇA_UNITARIA'] * df_ressarcimento['QUANTIDADE']

        total_geral = df_ressarcimento['VALOR_TOTAL_RESSARCIMENTO'].sum()
        st.metric("💵 Valor Total Geral de Ressarcimento", f"R$ {total_geral:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        
        st.markdown("---")
        colunas_res = [c for c in ['CODIGO_LIMPO', 'COD. EAN_LIMPO', 'NOME COMERCIAL', 'VALOR_RECOMENDADO_SP', 'PREÇO PRATICADO', 'QUANTIDADE', 'DIFERENÇA_UNITARIA', 'VALOR_TOTAL_RESSARCIMENTO'] if c in df_ressarcimento.columns]
        st.dataframe(
            df_ressarcimento[colunas_res],
            column_config={
                "CODIGO_LIMPO": "Código",
                "COD. EAN_LIMPO": "EAN",
                "NOME COMERCIAL": "Produto",
                "VALOR_RECOMENDADO_SP": st.column_config.NumberColumn("Rec. SP", format="R$ %.2f"),
                "PREÇO PRATICADO": st.column_config.NumberColumn("Praticado", format="R$ %.2f"),
                "QUANTIDADE": "Qtd",
                "DIFERENÇA_UNITARIA": st.column_config.NumberColumn("Dif. Unit.", format="R$ %.2f"),
                "VALOR_TOTAL_RESSARCIMENTO": st.column_config.NumberColumn("Total Apurado", format="R$ %.2f")
            },
            hide_index=True,
            use_container_width=True
        )

st.markdown("<br><hr><p style='text-align: center; color: #555555; font-size: 11px;'>Desenvolvido por Benedito Bandola</p>", unsafe_allow_html=True)
