import streamlit as st
import pandas as pd

df = pd.read_excel("TabelaEstoque.xlsx")

vendas = df[df["Tipo"] == "Venda"]
compras = df[df["Tipo"] == "Compra"]

qtd_vendida = vendas["Quantidade"].sum()
qtd_comprada = compras["Quantidade"].sum()

valor_vendido = (vendas["Quantidade"] * vendas["Preço Unitário"]).sum()
valor_comprado = (compras["Quantidade"] * compras["Preço Unitário"]).sum()

st.title("Controle de Estoque")

st.subheader("Vendas")
st.write("Quantidade vendida:", qtd_vendida)
st.write("Valor vendido: R$", valor_vendido)

st.subheader("Compras")
st.write("Quantidade comprada:", qtd_comprada)
st.write("Valor comprado: R$", valor_comprado)