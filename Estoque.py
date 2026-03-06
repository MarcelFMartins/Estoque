import streamlit as st
import pandas as pd
import altair as alt
from datetime import date
from openpyxl import load_workbook

st.set_page_config(layout="wide")

@st.dialog("Cadastrar Movimentação")
def cadastrar_produto():
    # O formulário dentro do modal
    data_mov = st.date_input("Data", value=date.today())
    produto_mov = st.selectbox("Produto", ["Morango", "Tomate"])
    tipo_mov = st.selectbox("Tipo", ["Compra", "Venda"])
    quantidade_mov = st.number_input("Quantidade", min_value=1)
    preco_mov = st.number_input("Preço Unitário", min_value=0.0)
    
    if st.button("Salvar Cadastro"):
        try:
            wb = load_workbook(arquivo)
            nome_da_aba = "Estoque" 
            ws = wb[nome_da_aba]
            
            proxima_linha = ws.max_row + 1
            
            # Fórmulas dinâmicas
            f_total = f"=(D{proxima_linha}*E{proxima_linha})"
            f_qtd = f'=IF(C{proxima_linha}="Compra"; G{proxima_linha-1}+D{proxima_linha}; G{proxima_linha-1}-D{proxima_linha})'
            f_custo = f'=IF(C{proxima_linha}="Compra"; (I{proxima_linha-1}+F{proxima_linha})/(G{proxima_linha}); H{proxima_linha-1})'
            f_val_estoque = f"=G{proxima_linha}*H{proxima_linha}"

            nova_linha = [data_mov, produto_mov, tipo_mov, quantidade_mov, preco_mov, f_total, f_qtd, f_custo, f_val_estoque]
            
            ws.append(nova_linha)
            ws.cell(row=proxima_linha, column=1).number_format = 'DD/MM/YYYY'
            
            wb.save(arquivo)
            st.success("Cadastrado com sucesso!")
            # st.rerun() serve para atualizar os gráficos e a tabela ao fundo imediatamente
            st.rerun() 
        except Exception as e:
            st.error(f"Erro ao salvar: {e}")

@st.dialog("Cadastrar Despesa")
def cadastrar_despesa():

    data_desp = st.date_input("Data", value=date.today())
    descricao = st.text_input("Descrição")
    categoria = st.selectbox("Categoria", ["Aluguel", "Fornecedor", "Salário", "Marketing", "Outros"])
    valor = st.number_input("Valor", min_value=0.0)

    if st.button("Salvar Despesa"):

        nova_despesa = pd.DataFrame({
            "Data": [data_desp],
            "Descrição": [descricao],
            "Categoria": [categoria],
            "Valor": [valor]
        })

        try:
            despesas = pd.read_excel("Despesas.xlsx")
            despesas = pd.concat([despesas, nova_despesa], ignore_index=True)
        except:
            despesas = nova_despesa

        despesas.to_excel("Despesas.xlsx", index=False)

        st.success("Despesa cadastrada!")
        st.rerun()

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

#BARRA LATERAL
if st.sidebar.button("➕ Cadastrar Produto"):
    cadastrar_produto()

if st.sidebar.button("💸 Cadastrar Despesa"):
    cadastrar_despesa()

#LOGOUT
if st.sidebar.button("🚪 Sair"):
    st.session_state["logado"] = False
    st.rerun()

#SISTEMA DE ESTOQUE
arquivo = "TabelaEstoque.xlsx"
df = pd.read_excel(arquivo)
df["Data"] = pd.to_datetime(df["Data"])
st.title("📦 Sistema Estoque")

#CALCULOS
vendas = df[df["Tipo"] == "Venda"]
compras = df[df["Tipo"] == "Compra"]
qtd_vendida = vendas["Quantidade"].sum()
qtd_comprada = compras["Quantidade"].sum()
valor_vendido = (vendas["Quantidade"] * vendas["Preço Unitário"]).sum()
valor_comprado = (compras["Quantidade"] * compras["Preço Unitário"]).sum()

#RESUMO
st.subheader("Resumo")
col1, col2, col3, col4, col5 = st.columns([2,2,2,2,2])
col1.metric("Quantidade Vendida", qtd_vendida)
col2.metric("Valor Vendido", f"R$ {valor_vendido:,.2f}")
col3.metric("Quantidade Comprada", qtd_comprada)
col4.metric("Valor Comprado", f"R$ {valor_comprado:,.2f}")
col5.metric("Lucro", f"R$ {(valor_vendido - valor_comprado):,.2f}")

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

#RESUMO VENDAS
st.title("💸 Resumo de Vendas por Produto")
vendas_produto = vendas.groupby("Produto").agg({
    "Quantidade": "sum",
    "Valor Total": "sum"
}).reset_index()

col1, col2 = st.columns(2)

with col1:
    st.subheader("Quantidade por Produto")
    grafico_qtd = alt.Chart(vendas_produto).mark_bar(color="red", size=30).encode(
        x=alt.X("Produto:N", title=""),
        y=alt.Y("Quantidade:Q", title="")
    ).properties(height=500)
    st.altair_chart(grafico_qtd, use_container_width=True)

with col2:
    st.subheader("Valor por Produto")
    grafico_valor = alt.Chart(vendas_produto).mark_bar(color="red", size=30).encode(
        x=alt.X("Produto:N", title=""),
        y=alt.Y("Valor Total:Q", title="")
    ).properties(height=500)
    st.altair_chart(grafico_valor, use_container_width=True)

st.title("📉 Despesas")

try:
    despesas = pd.read_excel("Despesas.xlsx")
    despesas["Data"] = pd.to_datetime(despesas["Data"])
except:
    despesas = pd.DataFrame(columns=["Data","Descrição","Categoria","Valor"])

st.dataframe(despesas, use_container_width=True)

st.subheader("Despesas por Categoria")

if not despesas.empty:

    despesas_categoria = despesas.groupby("Categoria")["Valor"].sum().reset_index()

    grafico_despesa = alt.Chart(despesas_categoria).mark_bar(
        color="orange",
        size=30
    ).encode(
        x=alt.X("Categoria:N", title=""),
        y=alt.Y("Valor:Q", title="Valor (R$)")
    ).properties(height=400)

    st.altair_chart(grafico_despesa, use_container_width=True)

st.title("📑 DRE")

receita_total = valor_vendido
cmv = valor_comprado
despesa_total = despesas["Valor"].sum() if not despesas.empty else 0

lucro_bruto = receita_total - cmv
lucro_liquido = lucro_bruto - despesa_total

col1, col2, col3, col4 = st.columns(4)

col1.metric("💰 Receita", f"R$ {receita_total:,.2f}")
col2.metric("📦 Custo das Compras", f"R$ {cmv:,.2f}")
col3.metric("📉 Despesas", f"R$ {despesa_total:,.2f}")
col4.metric("📈 Lucro Líquido", f"R$ {lucro_liquido:,.2f}")