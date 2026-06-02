import streamlit as st
import anthropic
import zipfile
import xml.etree.ElementTree as ET
import openpyxl
import os

# ── Configuração da página ──────────────────────────────────────────────────
st.set_page_config(
    page_title="Agente de Cartões Brasil 2026",
    page_icon="💳",
    layout="centered",
)

# ── Leitura dos arquivos de conhecimento ────────────────────────────────────
@st.cache_data(show_spinner="Carregando base de conhecimento...")
def carregar_base():
    base_dir = os.path.dirname(os.path.abspath(__file__))

    # .docx
    docx_path = os.path.join(base_dir, "referencia_cartoes_agente_2026.docx")
    paragrafos = []
    with zipfile.ZipFile(docx_path, "r") as z:
        with z.open("word/document.xml") as f:
            tree = ET.parse(f)
            ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
            for para in tree.findall(".//w:p", ns):
                texto = "".join(r.text or "" for r in para.findall(".//w:t", ns))
                if texto.strip():
                    paragrafos.append(texto.strip())
    doc_texto = "\n".join(paragrafos)

    # .xlsx
    xlsx_path = os.path.join(base_dir, "cartoes_credito_brasil_2026_v2.xlsx")
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    abas = []
    for nome in wb.sheetnames:
        ws = wb[nome]
        linhas = []
        for row in ws.iter_rows(values_only=True):
            linha = "\t".join(str(c) if c is not None else "" for c in row)
            if linha.strip():
                linhas.append(linha)
        abas.append(f"=== ABA: {nome} ===\n" + "\n".join(linhas))
    xlsx_texto = "\n\n".join(abas)

    sistema = f"""Você é o Agente de Cartões de Crédito Brasil 2026.
Sua base de conhecimento contém dados detalhados sobre cartões de crédito brasileiros, programas de pontos, perfis de usuário e estratégias de recomendação.

=== DOCUMENTO DE REFERÊNCIA ===
{doc_texto}

=== PLANILHA DE CARTÕES ===
{xlsx_texto}

Regras:
- Seja direto, claro e objetivo.
- Use formatação markdown nas respostas.
- Baseie todas as respostas EXCLUSIVAMENTE na base de conhecimento acima.
- Se uma informação não estiver na base, diga que não tem essa informação.
"""
    return sistema


def chamar_claude(sistema: str, mensagem: str) -> str:
    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        api_key = st.session_state.get("api_key", "")
    client = anthropic.Anthropic(api_key=api_key)

    resposta = ""
    with client.messages.stream(
        model="claude-sonnet-4-6",
        max_tokens=2048,
        system=[{"type": "text", "text": sistema, "cache_control": {"type": "ephemeral"}}],
        messages=[{"role": "user", "content": mensagem}],
    ) as stream:
        for texto in stream.text_stream:
            resposta += texto
    return resposta


# ── Estado da sessão ─────────────────────────────────────────────────────────
if "api_key" not in st.session_state:
    st.session_state.api_key = os.environ.get("ANTHROPIC_API_KEY", "")
if "modo" not in st.session_state:
    st.session_state.modo = None
if "resultado" not in st.session_state:
    st.session_state.resultado = ""

# ── Header ───────────────────────────────────────────────────────────────────
st.title("💳 Agente de Cartões de Crédito")
st.caption("Brasil 2026 · Powered by Claude")

# ── API Key (se não tiver no ambiente) ───────────────────────────────────────
if not st.session_state.api_key:
    with st.expander("🔑 Configurar chave de API", expanded=True):
        chave = st.text_input("ANTHROPIC_API_KEY", type="password", placeholder="sk-ant-...")
        if st.button("Salvar"):
            st.session_state.api_key = chave
            st.rerun()
    st.stop()

# ── Carrega base ─────────────────────────────────────────────────────────────
try:
    sistema = carregar_base()
except FileNotFoundError as e:
    st.error(f"Arquivo não encontrado: {e}\n\nCertifique-se de que `app.py` está na mesma pasta que os arquivos `.docx` e `.xlsx`.")
    st.stop()

# ── Menu principal ────────────────────────────────────────────────────────────
st.divider()
col1, col2, col3, col4 = st.columns(4)
with col1:
    if st.button("🎯 Recomendar", use_container_width=True):
        st.session_state.modo = "recomendar"
        st.session_state.resultado = ""
with col2:
    if st.button("⚖️ Comparar", use_container_width=True):
        st.session_state.modo = "comparar"
        st.session_state.resultado = ""
with col3:
    if st.button("💰 Vale a anuidade?", use_container_width=True):
        st.session_state.modo = "anuidade"
        st.session_state.resultado = ""
with col4:
    if st.button("📋 Explicar cartão", use_container_width=True):
        st.session_state.modo = "explicar"
        st.session_state.resultado = ""

st.divider()

# ── Modo: Recomendar ──────────────────────────────────────────────────────────
if st.session_state.modo == "recomendar":
    st.subheader("🎯 Recomendar cartão para o meu perfil")
    with st.form("form_recomendar"):
        renda = st.selectbox("Renda mensal aproximada", [
            "Até R$ 1.000", "R$ 1.000 – R$ 3.000", "R$ 3.000 – R$ 5.000",
            "R$ 5.000 – R$ 10.000", "R$ 10.000 – R$ 20.000", "Acima de R$ 20.000"
        ])

        st.markdown("**Objetivos** (selecione todos que se aplicam)")
        obj_latam   = st.checkbox("Milhas LATAM")
        obj_azul    = st.checkbox("Milhas Azul")
        obj_smiles  = st.checkbox("Milhas Smiles (GOL)")
        obj_cash    = st.checkbox("Cashback")
        obj_lounge  = st.checkbox("Lounge / Sala VIP")
        obj_gratis  = st.checkbox("Sem anuidade / custo zero")

        st.markdown("**Investimentos em bancos** (selecione todos que se aplicam)")
        col_a, col_b = st.columns(2)
        with col_a:
            inv_xp      = st.checkbox("XP Investimentos")
            inv_btg     = st.checkbox("BTG Pactual")
            inv_inter   = st.checkbox("Banco Inter")
            inv_bb      = st.checkbox("Banco do Brasil")
        with col_b:
            inv_brad    = st.checkbox("Bradesco")
            inv_sant    = st.checkbox("Santander")
            inv_itau    = st.checkbox("Itaú")
            inv_nenhum  = st.checkbox("Nenhum", value=True)

        st.markdown("**Valor investido por banco** (deixe em branco se não tiver)")
        col1i, col2i = st.columns(2)
        with col1i:
            val_xp    = st.number_input("XP (R$)",      min_value=0, step=10000, value=0)
            val_btg   = st.number_input("BTG (R$)",     min_value=0, step=10000, value=0)
            val_inter = st.number_input("Inter (R$)",   min_value=0, step=10000, value=0)
            val_bb    = st.number_input("BB (R$)",      min_value=0, step=10000, value=0)
        with col2i:
            val_brad  = st.number_input("Bradesco (R$)",  min_value=0, step=10000, value=0)
            val_sant  = st.number_input("Santander (R$)", min_value=0, step=10000, value=0)
            val_itau  = st.number_input("Itaú (R$)",      min_value=0, step=10000, value=0)

        viaja = st.radio("Viaja com frequência?", ["Sim", "Não"], horizontal=True)
        submitted = st.form_submit_button("Recomendar", type="primary", use_container_width=True)

    if submitted:
        objetivos = [o for o, v in [
            ("Milhas LATAM", obj_latam), ("Milhas Azul", obj_azul),
            ("Milhas Smiles (GOL)", obj_smiles), ("Cashback", obj_cash),
            ("Lounge / Sala VIP", obj_lounge), ("Sem anuidade / custo zero", obj_gratis)
        ] if v] or ["Não especificado"]

        investimentos = []
        for banco, marcado, valor in [
            ("XP Investimentos", inv_xp, val_xp), ("BTG Pactual", inv_btg, val_btg),
            ("Banco Inter", inv_inter, val_inter), ("Banco do Brasil", inv_bb, val_bb),
            ("Bradesco", inv_brad, val_brad), ("Santander", inv_sant, val_sant),
            ("Itaú", inv_itau, val_itau),
        ]:
            if marcado:
                investimentos.append(f"{banco}: R$ {valor:,.0f}" if valor > 0 else banco)
        inv_texto = ", ".join(investimentos) if investimentos else "Nenhum"

        prompt = (
            f"Perfil do usuário:\n"
            f"- Renda mensal: {renda}\n"
            f"- Objetivos: {', '.join(objetivos)}\n"
            f"- Investimentos: {inv_texto}\n"
            f"- Viaja frequentemente: {viaja}\n\n"
            f"Com base nesse perfil, recomende até 3 cartões de crédito. "
            f"Para cada um, explique por que faz sentido, qual a anuidade e se há algum alerta para 2026. "
            f"Leve em conta os investimentos bancários para identificar cartões exclusivos por relacionamento."
        )
        with st.spinner("Analisando seu perfil..."):
            st.session_state.resultado = chamar_claude(sistema, prompt)

# ── Modo: Comparar ────────────────────────────────────────────────────────────
elif st.session_state.modo == "comparar":
    st.subheader("⚖️ Comparar dois cartões")
    with st.form("form_comparar"):
        cartao1 = st.text_input("Primeiro cartão", placeholder="ex: Nubank Ultravioleta")
        cartao2 = st.text_input("Segundo cartão", placeholder="ex: C6 Carbon")
        submitted = st.form_submit_button("Comparar", type="primary", use_container_width=True)

    if submitted and cartao1 and cartao2:
        prompt = (
            f"Compare os cartões '{cartao1}' e '{cartao2}' lado a lado. "
            f"Use uma tabela markdown com os seguintes critérios: anuidade, programa de pontos, "
            f"taxa de conversão, benefícios principais, lounge, seguros, e para quem é mais indicado."
        )
        with st.spinner("Comparando cartões..."):
            st.session_state.resultado = chamar_claude(sistema, prompt)

# ── Modo: Vale a anuidade ─────────────────────────────────────────────────────
elif st.session_state.modo == "anuidade":
    st.subheader("💰 Vale a pena pagar a anuidade?")
    with st.form("form_anuidade"):
        cartao = st.text_input("Qual cartão?", placeholder="ex: Itaú Personnalité Visa Infinite")
        gasto_mensal = st.number_input("Gasto médio mensal no cartão (R$)", min_value=0, value=3000, step=500)
        uso_lounge = st.radio("Usa salas VIP em aeroportos?", ["Sim", "Não", "Às vezes"], horizontal=True)
        submitted = st.form_submit_button("Calcular", type="primary", use_container_width=True)

    if submitted and cartao:
        prompt = (
            f"O usuário quer saber se vale a pena pagar a anuidade do '{cartao}'.\n"
            f"- Gasto médio mensal: R$ {gasto_mensal:,.0f}\n"
            f"- Uso de sala VIP: {uso_lounge}\n\n"
            f"Faça o cálculo com números reais: quantos pontos/milhas ele acumula por ano, "
            f"quanto isso vale em reais, some o valor dos benefícios que ele usa, "
            f"e compare com o custo da anuidade. Conclua se vale ou não."
        )
        with st.spinner("Calculando..."):
            st.session_state.resultado = chamar_claude(sistema, prompt)

# ── Modo: Explicar ────────────────────────────────────────────────────────────
elif st.session_state.modo == "explicar":
    st.subheader("📋 Explicar um cartão específico")
    with st.form("form_explicar"):
        cartao = st.text_input("Nome do cartão", placeholder="ex: XP Visa Infinite")
        submitted = st.form_submit_button("Explicar", type="primary", use_container_width=True)

    if submitted and cartao:
        prompt = (
            f"Explique tudo sobre o cartão '{cartao}': anuidade, como isentar, programa de pontos, "
            f"taxa de conversão por categoria, benefícios (lounge, seguros, concierge), "
            f"renda mínima exigida, para qual perfil é indicado e alertas importantes para 2026."
        )
        with st.spinner("Buscando informações..."):
            st.session_state.resultado = chamar_claude(sistema, prompt)

# ── Resultado ─────────────────────────────────────────────────────────────────
if st.session_state.resultado:
    st.divider()
    st.markdown(st.session_state.resultado)
    if st.button("🔄 Nova consulta"):
        st.session_state.resultado = ""
        st.session_state.modo = None
        st.rerun()
