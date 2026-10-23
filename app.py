import os
import unicodedata
import pandas as pd
import streamlit as st

# --- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="Auditoria e Ressarcimento - São Paulo",
    layout="wide",
    page_icon="📊",
)

# --- 2. DICIONÁRIO DE LOJAS POR FILIAL ---
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
        # Adicione aqui as lojas de Ribeirão Preto seguindo o mesmo formato quando necessário
        {"codigo": 1001, "nome": "EXEMPLO LOJA RIBEIRÃO PRETO LTDA", "cnpj": "00.000.000/0001-00", "cidade": "RIBEIRAO PRETO", "bairro": "CENTRO"}
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
.card-relatorio {
    background: #FFFFFF;
    padding: 20px;
    border-radius: 12px;
    border: 1px solid #E2E8F0;
    box-shadow: 0 4px 6px rgba(0,0,0,0.02);
    margin-bottom: 15px;
}
</style>
""",
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="cabecalho-principal"><h2 style="color: #0F172A; margin:0;">📊 Sistema de Auditoria e Ressarcimento (Estado de SP)</h2><p style="color: #64748B; font-size: 13px; margin:5px 0 0 0;">Controle de preços praticados, distorções de SKU e cálculo de ressarcimento por loja.</p></div>',
    unsafe_allow_html=True,
)

# --- 4. FUNÇÕES DE SUPORTE E CARREGAMENTO DE DADOS ---
@st.cache_data
def carregar_dados():
    arquivo_base = "boneco_com_valor_h.xlsx"
    if not os.path.exists(arquivo_base):
        return None
    try:
        df = pd.read_excel(arquivo_base, sheet_name="Detalhado")
    except Exception:
        try:
            df = pd.read_excel(arquivo_base)
        except Exception as e:
            st.error(f"Erro ao abrir a planilha: {e}")
            return None

    df.columns = [str(c).sri().upper() if hasattr(str(c), 'sri') else str(c).strip().upper() for c in df.columns]
    
    # Limpeza de campos de código
    for col in ["CODIGO", "COD. EAN", "SKU"]:
        if col in df.columns:
            df[f"{col}_LIMPO"] = df[col].astype(str).str.replace(".0", "", regex=False).str.strip()
        else:
            df[f"{col}_LIMPO"] = ""
            
    return df

df_produtos = carregar_dados()

if df_produtos is None:
    st.error("⚠️ O arquivo `boneco_com_valor_h.xlsx` não foi encontrado na raiz do projeto. Por favor, envie-o para o repositório.")
    st.stop()

# --- 5. SELEÇÃO DE FILIAL E LOJA (BARRA LATERAL) ---
st.sidebar.header("📍 Configuração da Auditoria")
filial_selecionada = st.sidebar.selectbox("Selecione a Filial:", ["SÃO JOSÉ DO RIO PRETO", "RIBEIRÃO PRETO"])

lista_lojas = LOJAS_POR_FILIAL.get(filial_selecionada, [])
loja_nomes = [f"{l['codigo']} - {l['nome']} ({l['cidade']})" for l in lista_lojas]

loja_escolhida_str = st.sidebar.selectbox("Selecione a Loja:", loja_nomes)

# Identificar o dicionário da loja selecionada
loja_obj = None
if loja_escolhida_str:
    codigo_loja_sel = int(loja_escolhida_str.split(" - ")[0])
    loja_obj = next((l for l in lista_lojas if l['codigo'] == codigo_loja_sel), None)

if loja_obj:
    st.sidebar.markdown("---")
    st.sidebar.markdown(f"**Loja:** {loja_obj['nome']}")
    st.sidebar.markdown(f"**CNPJ:** {loja_obj['cnpj']}")
    st.sidebar.markdown(f"**Cidade/Bairro:** {loja_obj['cidade']} / {loja_obj['bairro']}")

# Inicializar o estado da sessão para armazenar os preços cadastrados por loja
if "auditoria_dados" not in st.session_state:
    st.session_state["auditoria_dados"] = {}

chave_sessao_loja = f"loja_{loja_obj['codigo']}" if loja_obj else "loja_padrao"

if chave_sessao_loja not in st.session_state["auditoria_dados"]:
    # Inicializa DataFrame de auditoria com os produtos da base e colunas de preço praticado e quantidade
    df_init = df_produtos[['CODIGO_LIMPO', 'NOME COMERCIAL', 'DESCRICAO', 'FAMILIA', 'VALOR_RECOMENDADO_SP']].copy()
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
        "1. Cadastrar / Atualizar Preços e Quantidades da Loja",
        "2. Relatório de Distorção de Preços (Apenas Preços)",
        "3. Relatório de Cálculo de Ressarcimento (Com Quantidades)"
    ],
    horizontal=True
)

st.markdown("---")

# --- MODO 1: CADASTRO DE PREÇOS E QUANTIDADES ---
if modo_operacao.startswith("1"):
    st.markdown(f"#### 📝 Cadastro de Preços - {loja_obj['nome'] if loja_obj else ''}")
    st.markdown("Insira o preço praticado pela loja e, opcionalmente, a quantidade vendida/em estoque para apuração de ressarcimento.")

    # Filtro rápido por família ou nome
    filtro_busca = st.text_input("Filtrar produto por nome ou código para agilizar o lançamento:", placeholder="Ex: GASTRO, KITTEN, 214731...")
    
    df_filtrado = df_loja_atual.copy()
    if filtro_busca:
        termo = filtro_busca.lower()
        df_filtrado = df_filtrado[
            df_filtrado['NOME COMERCIAL'].astype(str).str.lower().str.contains(termo) |
            df_filtrado['DESCRICAO'].astype(str).str.lower().str.contains(termo) |
            df_filtrado['CODIGO_LIMPO'].astype(str).str.contains(termo)
        ]

    # Usando st.data_editor para edição interativa em tabela
    colunas_edicao = ['CODIGO_LIMPO', 'NOME COMERCIAL', 'FAMILIA', 'VALOR_RECOMENDADO_SP', 'PREÇO PRATICADO', 'QUANTIDADE']
    
    editado_df = st.data_editor(
        df_filtrado[colunas_edicao],
        column_config={
            "CODIGO_LIMPO": st.column_config.TextColumn("Código", disabled=True),
            "NOME COMERCIAL": st.column_config.TextColumn("Produto", disabled=True),
            "FAMILIA": st.column_config.TextColumn("Família", disabled=True),
            "VALOR_RECOMENDADO_SP": st.column_config.NumberColumn("Preço Rec. (SP)", format="R$ %.2f", disabled=True),
            "PREÇO PRATICADO": st.column_config.NumberColumn("Preço Praticado na Loja (R$)", format="R$ %.2f", min_value=0.0, step=0.1),
            "QUANTIDADE": st.column_config.NumberColumn("Qtd (para cálculo)", min_value=0, step=1)
        },
        hide_index=True,
        use_container_width=True,
        key=f"editor_{chave_sessao_loja}"
    )

    if st.button("💾 Salvar Alterações desta Loja", type="primary"):
        # Atualizar no estado global da sessão
        for idx, row in editado_df.iterrows():
            codigo_prod = row['CODIGO_LIMPO']
            mask = st.session_state["auditoria_dados"][chave_sessao_loja]['CODIGO_LIMPO'] == codigo_prod
            st.session_state["auditoria_dados"][chave_sessao_loja].loc[mask, 'PREÇO PRATICADO'] = row['PREÇO PRATICADO']
            st.session_state["auditoria_dados"][chave_sessao_loja].loc[mask, 'QUANTIDADE'] = row['QUANTIDADE']
            if row['PREÇO PRATICADO'] > 0:
                st.session_state["auditoria_dados"][chave_sessao_loja].loc[mask, 'STATUS_AUDITORIA'] = "Auditado"
                
        st.success("✅ Status e preços salvos com sucesso para esta loja!")

# --- MODO 2: RELATÓRIO DE DISTORÇÃO DE PREÇOS ---
elif modo_operacao.startswith("2"):
    st.markdown(f"#### 🔍 Relatório de Distorção de Preços (Apenas Preços)")
    st.markdown("Mostra a diferença unitária por SKU entre o preço praticado na loja e o recomendado para o estado de São Paulo.")

    df_atual = st.session_state["auditoria_dados"][chave_sessao_loja]
    df_auditado = df_atual[df_atual['PREÇO PRATICADO'] > 0].copy()

    if df_auditado.empty:
        st.warning("⚠️ Nenhum produto foi auditado com preço praticado ainda. Vá para o Modo 1 para cadastrar.")
    else:
        df_auditado['DIFERENÇA_UNITARIA'] = df_auditado['PREÇO PRATICADO'] - df_auditado['VALOR_RECOMENDADO_SP']
        df_auditado['TIPO_DISTORCAO'] = df_auditado['DIFERENÇA_UNITARIA'].apply(lambda x: "Acima do Recomendado" if x > 0 else ("Abaixo do Recomendado" if x < 0 else "Preço Exato"))

        # Exibir métricas gerais
        total_skus = len(df_auditado)
        skus_com_distorcao = len(df_auditado[df_auditado['DIFERENÇA_UNITARIA'] != 0])
        
        m1, m2 = st.columns(2)
        m1.metric("SKUs Auditados", total_skus)
        m2.metric("SKUs com Diferença de Preço", skus_com_distorcao)

        st.markdown("---")
        
        # Tabela formatada de distorção
        tabela_distorcao = df_auditado[['CODIGO_LIMPO', 'NOME COMERCIAL', 'FAMILIA', 'VALOR_RECOMENDADO_SP', 'PREÇO PRATICADO', 'DIFERENÇA_UNITARIA', 'TIPO_DISTORCAO']].copy()
        
        st.dataframe(
            tabela_distorcao,
            column_config={
                "CODIGO_LIMPO": "Código",
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
    st.markdown("Calcula o valor total a ressarcir por SKU com base na diferença de preço multiplicada pela quantidade informada.")

    df_atual = st.session_state["auditoria_dados"][chave_sessao_loja]
    df_ressarcimento = df_atual[(df_atual['PREÇO PRATICADO'] > 0) & (df_atual['QUANTIDADE'] > 0)].copy()

    if df_ressarcimento.empty:
        st.warning("⚠️ Nenhum produto possui preço praticado E quantidade cadastrados simultaneamente. Preencha no Modo 1.")
    else:
        # Diferença unitária (se o praticado for menor que o recomendado, ou regra específica de ressarcimento)
        # Aqui calculamos a diferença absoluta ou conforme regra de negócio (Praticado vs Recomendado) * Qtd
        df_ressarcimento['DIFERENÇA_UNITARIA'] = df_ressarcimento['PREÇO PRATICADO'] - df_ressarcimento['VALOR_RECOMENDADO_SP']
        df_ressarcimento['VALOR_TOTAL_RESSARCIMENTO'] = df_ressarcimento['DIFERENÇA_UNITARIA'] * df_ressarcimento['QUANTIDADE']

        total_geral_ressarcimento = df_ressarcimento['VALOR_TOTAL_RESSARCIMENTO'].sum()

        st.metric("💵 Valor Total Geral de Apuração / Ressarcimento", f"R$ {total_geral_ressarcimento:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        
        st.markdown("---")

        tabela_ressarcimento = df_ressarcimento[['CODIGO_LIMPO', 'NOME COMERCIAL', 'VALOR_RECOMENDADO_SP', 'PREÇO PRATICADO', 'QUANTIDADE', 'DIFERENÇA_UNITARIA', 'VALOR_TOTAL_RESSARCIMENTO']].copy()

        st.dataframe(
            tabela_ressarcimento,
            column_config={
                "CODIGO_LIMPO": "Código",
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

# Assinatura de autoria
st.markdown("<br><hr><p style='text-align: center; color: #555555; font-size: 11px;'>Desenvolvido por Benedito Bandola</p>", unsafe_allow_html=True)
