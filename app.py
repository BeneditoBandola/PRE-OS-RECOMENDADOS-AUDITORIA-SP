import os
json = __import__('json')
unicodedata = __import__('unicodedata')
pandas = __import__('pandas')
streamlit = st = __import__('streamlit')

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT

# --- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="Ajuste de Preços - São Paulo",
    layout="wide",
    page_icon="📊",
)

# --- 2. DICIONÁRIO DE LOJAS REAIS (FILIAIS DE SP) ---
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
    "RIBEIRÃO PRETO (FILIAL SÃO JOÃO)": [
        {"codigo": 14272, "nome": "AMOPETS LTDA", "cnpj": "13.062.215/0001-89", "cidade": "RIBEIRAO PRETO", "bairro": "JARDIM IRAJA"},
        {"codigo": 13829, "nome": "CENTER SHOP DO ANIMAL LTDA", "cnpj": "30.872.481/0001-96", "cidade": "RIBEIRAO PRETO", "bairro": "BONFIM PAULISTA"},
        {"codigo": 10421, "nome": "DELIBO PET SHOP LTDA", "cnpj": "17.974.159/0001-00", "cidade": "RIBEIRAO PRETO", "bairro": "VILA TIBERIO"},
        {"codigo": 16261, "nome": "E.E.S. PET SHOP COM DE PRODS E ANIMAIS DE ESTIM LTDA", "cnpj": "36.667.802/0001-05", "cidade": "RIBEIRAO PRETO", "bairro": "PLANALTO VERDE"},
        {"codigo": 24050, "nome": "JB PETSHOP RP LTDA", "cnpj": "52.806.664/0001-46", "cidade": "RIBEIRAO PRETO", "bairro": "SANTA CRUZ DO JOSE JACQUES"},
        {"codigo": 8312, "nome": "TCM AGROPET LTDA", "cnpj": "69.061.695/0001-18", "cidade": "RIBEIRAO PRETO", "bairro": "VILA VIRGINIA"},
        {"codigo": 8785, "nome": "BIVETER COMERCIO DE PRODUTOS AGROPECUARIOS LTDA", "cnpj": "07.831.582/0001-05", "cidade": "SAO CARLOS", "bairro": "JD. RICETTI"},
        {"codigo": 24457, "nome": "CASA PET C LTDA", "cnpj": "56.961.304/0001-70", "cidade": "SAO CARLOS", "bairro": "CENTRO"},
        {"codigo": 628, "nome": "COMAV COMERCIAL AGRO PECUARIA LTDA", "cnpj": "59.723.817/0001-30", "cidade": "SAO CARLOS", "bairro": "CENTRO"},
        {"codigo": 23333, "nome": "PULL PET LTDA", "cnpj": "49.703.428/0001-90", "cidade": "SAO CARLOS", "bairro": "CENTRO"}
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
.cabecalho-principal h2 { color: #0F172A !important; }
.cabecalho-principal p { color: #475569 !important; }
</style>
""",
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="cabecalho-principal"><h2>📊 Ajuste de Preços</h2><p>Busca rápida por EAN/Código, controle de preços, salvamento em JSON e Exportação automática em PDF.</p></div>',
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
filial_selecionada = st.sidebar.selectbox("Selecione a Filial:", ["SÃO JOSÉ DO RIO PRETO", "RIBEIRÃO PRETO (FILIAL SÃO JOÃO)"])

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
            for c in ['CODIGO_LIMPO', 'COD. EAN_LIMPO', 'SKU_LIMPO']:
                if c in df_carregado.columns:
                    df_carregado[c] = df_carregado[c].astype(str).str.replace(".0", "", regex=False).str.strip()
        except Exception:
            df_carregado = None

    if df_carregado is not None and not df_carregado.empty:
        st.session_state["auditoria_dados"][chave_sessao_loja] = df_carregado
    else:
        colunas_necessarias = [c for c in ['CODIGO_LIMPO', 'COD. EAN_LIMPO', 'SKU_LIMPO', 'NOME COMERCIAL', 'FAMILIA', 'VALOR_RECOMENDADO_SP'] if c in df_produtos.columns]
        df_init = df_produtos[colunas_necessarias].copy()
        df_init['PREÇO PRATICADO'] = 0.0
        df_init['QUANTIDADE'] = 0
        df_init['STATUS_AUDITORIA'] = "Pendente"
        st.session_state["auditoria_dados"][chave_sessao_loja] = df_init

df_loja_atual = st.session_state["auditoria_dados"][chave_sessao_loja]

# --- 6. ESCOLHA DO MODO DE OPERAÇÃO (BOTÕES GRANDES) ---
st.markdown("### 🛠️ Escolha a Opção Desejada")

col_b1, col_b2 = st.columns(2)

with col_b1:
    btn_opcao_1 = st.button("📝 1. Cadastrar Preços, Quantidades & Salvar", use_container_width=True, type="primary")
with col_b2:
    btn_opcao_2 = st.button("🔍 2. Listar Somente Preços Diferentes (Distorção)", use_container_width=True)

if "modo_ativo" not in st.session_state:
    st.session_state["modo_ativo"] = "cadastro"

if btn_opcao_1:
    st.session_state["modo_ativo"] = "cadastro"
if btn_opcao_2:
    st.session_state["modo_ativo"] = "relatorio"

st.markdown("---")

# Função de geração de PDF
def gerar_pdf_ressarcimento(loja, df_dados, total):
    pdf_path = f"Relatorio_Ressarcimento_{loja['codigo']}.pdf"
    doc = SimpleDocTemplate(pdf_path, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=14, textColor=colors.HexColor("#C0392B"), alignment=TA_CENTER)
    sub_style = ParagraphStyle('SubStyle', parent=styles['Normal'], fontName='Helvetica', fontSize=10, textColor=colors.HexColor("#2C3E50"), alignment=TA_CENTER)
    cell_style = ParagraphStyle('CellStyle', parent=styles['Normal'], fontName='Helvetica', fontSize=8.5, textColor=colors.HexColor("#2C3E50"))
    cell_header = ParagraphStyle('CellHeader', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=9, textColor=colors.white, alignment=TA_CENTER)

    story = []
    story.append(Paragraph("<b>RELATÓRIO DE AJUSTE DE PREÇOS E RESSARCIMENTO</b>", title_style))
    story.append(Spacer(1, 5))
    story.append(Paragraph(f"<b>Loja:</b> {loja['nome']} | <b>CNPJ:</b> {loja['cnpj']} | <b>Cidade:</b> {loja['cidade']}", sub_style))
    story.append(Spacer(1, 15))

    tabela_dados = [[
        Paragraph("<b>Cód</b>", cell_header),
        Paragraph("<b>Produto</b>", cell_header),
        Paragraph("<b>Rec. SP</b>", cell_header),
        Paragraph("<b>Praticado</b>", cell_header),
        Paragraph("<b>Qtd</b>", cell_header),
        Paragraph("<b>Dif. Unit.</b>", cell_header),
        Paragraph("<b>Total Apurado</b>", cell_header)
    ]]

    for _, row in df_dados.iterrows():
        tabela_dados.append([
            Paragraph(str(row['CODIGO_LIMPO']), cell_style),
            Paragraph(str(row['NOME COMERCIAL']), cell_style),
            Paragraph(f"R$ {row['VALOR_RECOMENDADO_SP']:.2f}", cell_style),
            Paragraph(f"R$ {row['PREÇO PRATICADO']:.2f}", cell_style),
            Paragraph(str(int(row['QUANTIDADE'])), cell_style),
            Paragraph(f"R$ {row['DIFERENÇA_UNITARIA']:.2f}", cell_style),
            Paragraph(f"R$ {row['VALOR_TOTAL_RESSARCIMENTO']:.2f}", cell_style)
        ])

    tabela = Table(tabela_dados, colWidths=[45, 160, 60, 60, 35, 60, 85])
    tabela.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#C0392B")),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#BDC3C7")),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
    ]))

    story.append(tabela)
    story.append(Spacer(1, 15))
    story.append(Paragraph(f"<b>VALOR TOTAL GERAL DE RESSARCIMENTO: R$ {total:,.2f}</b>".replace(",", "X").replace(".", ",").replace("X", "."), ParagraphStyle('TotalStyle', parent=styles['Heading2'], fontName='Helvetica-Bold', fontSize=12, textColor=colors.HexColor("#27AE60"), alignment=TA_CENTER)))

    doc.build(story)
    return pdf_path

# --- OPÇÃO 1: CADASTRAR PREÇOS, QUANTIDADES E SALVAR (COM PDF AUTOMÁTICO) ---
if st.session_state["modo_ativo"] == "cadastro":
    st.markdown(f"#### 📝 Cadastro de Preços e Quantidades - {loja_obj['nome'] if loja_obj else ''}")
    st.markdown("💡 **Dica:** Digite os **últimos números do EAN**, código ou nome para localizar o produto instantaneamente.")

    filtro_busca = st.text_input("🔍 Pesquisa rápida (EAN, Código ou Nome):", placeholder="Ex: 4731, gastro, kitten...")
    
    df_filtrado = df_loja_atual.copy()
    if filtro_busca:
        termo = filtro_busca.strip().lower()
        
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

    if st.button("💾 Salvar em JSON & Baixar PDF Oficial", type="primary"):
        for idx, row in editado_df.iterrows():
            codigo_prod = row['CODIGO_LIMPO']
            mask = st.session_state["auditoria_dados"][chave_sessao_loja]['CODIGO_LIMPO'] == codigo_prod
            st.session_state["auditoria_dados"][chave_sessao_loja].loc[mask, 'PREÇO PRATICADO'] = row['PREÇO PRATICADO']
            st.session_state["auditoria_dados"][chave_sessao_loja].loc[mask, 'QUANTIDADE'] = row['QUANTIDADE']
            if row['PREÇO PRATICADO'] > 0:
                st.session_state["auditoria_dados"][chave_sessao_loja].loc[mask, 'STATUS_AUDITORIA'] = "Auditado"
        
        # Salva no arquivo JSON
        df_para_salvar = st.session_state["auditoria_dados"][chave_sessao_loja]
        df_para_salvar.to_json(caminho_json, orient="records", force_ascii=False, indent=4)
        
        st.success(f"✅ Dados salvos com sucesso e gravados no arquivo `{caminho_json}`!")

        # Prepara e disponibiliza o PDF para download direto na tela
        df_atual = st.session_state["auditoria_dados"][chave_sessao_loja]
        df_ressarcimento = df_atual[(df_atual['PREÇO PRATICADO'] > 0) & (df_atual['QUANTIDADE'] > 0)].copy()

        if not df_ressarcimento.empty:
            df_ressarcimento['DIFERENÇA_UNITARIA'] = df_ressarcimento['PREÇO PRATICADO'] - df_ressarcimento['VALOR_RECOMENDADO_SP']
            df_ressarcimento['VALOR_TOTAL_RESSARCIMENTO'] = df_ressarcimento['DIFERENÇA_UNITARIA'] * df_ressarcimento['QUANTIDADE']
            total_geral = df_ressarcimento['VALOR_TOTAL_RESSARCIMENTO'].sum()

            arquivo_pdf = gerar_pdf_ressarcimento(loja_obj, df_ressarcimento, total_geral)
            with open(arquivo_pdf, "rb") as pdf_file:
                st.download_button(
                    label="⬇️ BAIXAR PDF DO RESSARCIMENTO AGORA",
                    data=pdf_file,
                    file_name=arquivo_pdf,
                    mime="application/pdf",
                    type="primary"
                )
        else:
            st.info("ℹ️ Nenhum item possui preço e quantidade maiores que zero para compor o PDF de ressarcimento.")

# --- OPÇÃO 2: LISTAR SOMENTE PREÇOS DIFERENTES (DISTORÇÃO) ---
else:
    st.markdown(f"#### 🔍 Listagem de Preços Diferentes (Distorção) - {loja_obj['nome'] if loja_obj else ''}")
    df_atual = st.session_state["auditoria_dados"][chave_sessao_loja]
    df_auditado = df_atual[df_atual['PREÇO PRATICADO'] > 0].copy()

    if df_auditado.empty:
        st.warning("⚠️ Nenhum produto foi cadastrado com preço praticado ainda.")
    else:
        df_auditado['DIFERENÇA_UNITARIA'] = df_auditado['PREÇO PRATICADO'] - df_auditado['VALOR_RECOMENDADO_SP']
        df_diferentes = df_auditado[df_auditado['DIFERENÇA_UNITARIA'] != 0].copy()

        if df_diferentes.empty:
            st.success("🎉 Todos os produtos auditados estão com o preço exato recomendado!")
        else:
            df_diferentes['TIPO_DISTORCAO'] = df_diferentes['DIFERENÇA_UNITARIA'].apply(lambda x: "Acima do Recomendado" if x > 0 else "Abaixo do Recomendado")

            st.metric("SKUs com Preços Diferentes", len(df_diferentes))
            st.markdown("---")
            
            colunas_rel = [c for c in ['CODIGO_LIMPO', 'COD. EAN_LIMPO', 'NOME COMERCIAL', 'FAMILIA', 'VALOR_RECOMENDADO_SP', 'PREÇO PRATICADO', 'DIFERENÇA_UNITARIA', 'TIPO_DISTORCAO'] if c in df_diferentes.columns]
            st.dataframe(
                df_diferentes[colunas_rel],
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

st.markdown("<br><hr><p style='text-align: center; color: #555555; font-size: 11px;'>Desenvolvido por Benedito Bandola</p>", unsafe_allow_html=True)
