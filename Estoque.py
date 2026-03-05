import streamlit as st
import pandas as pd
from datetime import date
from openpyxl import load_workbook

st.set_page_config(layout="wide")

#USUÁRIOS
usuarios = {
    "admin": "1234",
    "marcel": "1234"
}

#LOGIN
def login():
    st.title("🔐 Login")
    usuario = st.text_input("Usuário")
    senha = st.text_input("Senha", type="password")
    if st.button("Entrar"):
        if usuario in usuarios and usuarios[usuario] == senha:
            st.session_state["logado"] = True
            st.rerun()
        else:
            st.error("Usuário ou senha incorretos")

#SESSÃO
if "logado" not in st.session_state:
    st.session_state["logado"] = False

if not st.session_state["logado"]:
    login()
    st.stop()

#LOGOUT
if st.sidebar.button("🚪 Sair"):
    st.session_state["logado"] = False
    st.rerun()

#SISTEMA DE ESTOQUE
arquivo = "TabelaEstoque.xlsx"
df = pd.read_excel(arquivo)
df["Data"] = pd.to_datetime(df["Data"])
st.title("📦 Sistema Estoque")

#CADASTRO
st.subheader("Cadastrar movimentação")

with st.form("movimentacao"):
    data = st.date_input("Data", value=date.today())
    produto = st.selectbox("Produto", ["Morango", "Tomate"])
    tipo = st.selectbox("Tipo", ["Compra", "Venda"])
    quantidade = st.number_input("Quantidade", min_value=1)
    preco = st.number_input("Preço Unitário", min_value=0.0)
    enviar = st.form_submit_button("Cadastrar")
    if enviar:
        wb = load_workbook(arquivo)
        ws = wb.active
        nova_linha = [
            data,
            produto,
            tipo,
            quantidade,
            preco
        ]
        ws.append(nova_linha)
        wb.save(arquivo)
        st.success("Movimentação cadastrada com sucesso!")

#CALCULOS
vendas = df[df["Tipo"] == "Venda"]
compras = df[df["Tipo"] == "Compra"]
qtd_vendida = vendas["Quantidade"].sum()
qtd_comprada = compras["Quantidade"].sum()
valor_vendido = (vendas["Quantidade"] * vendas["Preço Unitário"]).sum()
valor_comprado = (compras["Quantidade"] * compras["Preço Unitário"]).sum()

#RESUMO
st.subheader("Resumo")
col1, col2, col3, col4 = st.columns([2,2,2,2])
col1.metric("Quantidade Vendida", qtd_vendida)
col2.metric("Valor Vendido", f"R$ {valor_vendido:,.2f}")
col3.metric("Quantidade Comprada", qtd_comprada)
col4.metric("Valor Comprado", f"R$ {valor_comprado:,.2f}")

#FUNÇÃO MOEDA BR
def moeda_br(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

#TABELA
df_view = df.copy()
df_view["Preço Unitário"] = df_view["Preço Unitário"].apply(moeda_br)
df_view["Valor Total"] = df_view["Valor Total"].apply(moeda_br)
df_view["Custo Médio"] = df_view["Custo Médio"].apply(moeda_br)
df_view["Valor Total em Estoque"] = df_view["Valor Total em Estoque"].apply(moeda_br)
st.dataframe(
    df_view,
    column_config={
        "Data": st.column_config.DateColumn(
            "Data",
            format="DD/MM/YYYY"
        ),
        "Quantidade": st.column_config.NumberColumn(
            "Quantidade",
            format="%d"
        )
    },
    use_container_width=True
)

#VENDAS POR PRODUTO
st.subheader("Vendas por Produto")
vendas_produto = vendas.groupby("Produto")["Quantidade"].sum()
st.bar_chart(vendas_produto)