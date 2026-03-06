import streamlit as st
import streamlit_authenticator as stauth

# USUÁRIOS
credentials = {
    "usernames": {
        "admin": {
            "name": "Administrador",
            "password": "1234"
        },
        "marcel": {
            "name": "Marcel",
            "password": "1234"
        },
        "matheus": {
            "name": "Matheus",
            "password": "1234"
        }
    }
}

authenticator = stauth.Authenticate(
    credentials,
    "estoque_cookie",
    "abcdef123456789",
    cookie_expiry_days=30
)

# MOSTRAR LOGIN
authenticator.login(location="main")


# PEGAR STATUS
authentication_status = st.session_state.get("authentication_status")
name = st.session_state.get("name")
username = st.session_state.get("username")

if authentication_status is False:
    st.error("Usuário ou senha incorretos")
    st.stop()

if authentication_status is None:
    st.warning("Digite usuário e senha")
    st.stop()

import pandas as pd
import altair as alt
from datetime import date
from openpyxl import load_workbook

st.set_page_config(layout="wide")
arquivo = "TabelaEstoque.xlsx"


# FUNÇÃO MOEDA BR
def moeda_br(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


# CADASTRAR MOVIMENTAÇÃO
@st.dialog("Cadastrar Movimentação")
def cadastrar_produto():

    data_mov = st.date_input("Data", value=date.today())

    try:
        df = pd.read_excel(arquivo)

        df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
        df["Produto"] = df["Produto"].astype(str)
        df["Tipo"] = df["Tipo"].astype(str)

        produtos_cadastrados = df["Produto"].dropna().unique().tolist()

    except:
        df = pd.DataFrame(columns=[
            "Data","Produto","Tipo","Quantidade","Preço Unitário",
            "Valor Total","Quantidade em Estoque","Custo Médio","Valor Total em Estoque"
        ])
        produtos_cadastrados = ["Morango","Tomate"]

    produto_mov = st.selectbox("Produto", produtos_cadastrados)
    tipo_mov = st.selectbox("Tipo", ["Compra","Venda"])
    quantidade_mov = st.number_input("Quantidade", min_value=1)
    preco_mov = st.number_input("Preço Unitário", min_value=0.0, format="%.2f")

    if st.button("Salvar Cadastro"):
        if tipo_mov == "Compra":
            contato_valor = "Fornecedor"
        else:
            contato_valor = "Cliente"
        try:

            valor_total = quantidade_mov * preco_mov

            nova_linha = pd.DataFrame([{
                "Data": pd.to_datetime(data_mov),
                "Produto": produto_mov,
                "Tipo": tipo_mov,
                "Cliente/Fornecedor": contato_valor,
                "Quantidade": quantidade_mov,
                "Preço Unitário": preco_mov,
                "Valor Total": valor_total
            }])

            df = pd.concat([df, nova_linha], ignore_index=True)

            df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
            df["Produto"] = df["Produto"].astype(str)

            df = df.sort_values(by=["Produto","Data"], ascending=[True,True]).reset_index(drop=True)

            df["Quantidade em Estoque"] = 0.0
            df["Custo Médio"] = 0.0
            df["Valor Total em Estoque"] = 0.0

            estoques = {}
            custos = {}

            for i, row in df.iterrows():

                prod = row["Produto"]
                tipo = row["Tipo"]
                qtd = float(row["Quantidade"])
                valor = float(row["Valor Total"])

                estoque_ant = estoques.get(prod,0)
                custo_ant = custos.get(prod,0)

                if tipo == "Compra":

                    novo_estoque = estoque_ant + qtd

                    custo_medio = (
                        (custo_ant * estoque_ant + valor) / novo_estoque
                        if novo_estoque != 0 else 0
                    )

                else:

                    novo_estoque = estoque_ant - qtd
                    custo_medio = custo_ant

                estoques[prod] = novo_estoque
                custos[prod] = custo_medio

                df.at[i,"Quantidade em Estoque"] = novo_estoque
                df.at[i,"Custo Médio"] = custo_medio
                df.at[i,"Valor Total em Estoque"] = novo_estoque * custo_medio

            df.to_excel(arquivo, index=False)

            st.success("Produto cadastrado com sucesso!")
            st.rerun()

        except Exception as e:
            st.error(f"Erro ao salvar: {e}")

# CADASTRAR DESPESA
@st.dialog("Cadastrar Despesa")
def cadastrar_despesa():

    data_desp = st.date_input("Data", value=date.today())
    descricao = st.text_input("Descrição")
    categoria = st.selectbox("Categoria", ["Combustível", "Alimentação", "Hotel", "Despesas com Funcionários", "Outros"])
    valor = st.number_input("Valor", min_value=0.0)

    if st.button("Salvar Despesa"):

        nova_despesa = pd.DataFrame({
            "Data": [data_desp],
            "Descrição": [descricao],
            "Fornecedor": ["Fornecedor"],
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

def recalcular_estoque(df):

    df = df.sort_values(["Produto","Data"]).reset_index(drop=True)

    df["Quantidade em Estoque"] = 0
    df["Custo Médio"] = 0
    df["Valor Total em Estoque"] = 0

    estoques = {}
    custos = {}

    for i, row in df.iterrows():

        prod = row["Produto"]
        tipo = row["Tipo"]
        qtd = row["Quantidade"]
        valor = row["Valor Total"]

        estoque_ant = estoques.get(prod,0)
        custo_ant = custos.get(prod,0)

        if tipo == "Compra":

            novo_estoque = estoque_ant + qtd

            custo_medio = (
                (custo_ant * estoque_ant + valor) / novo_estoque
                if novo_estoque != 0 else 0
            )

        else:

            novo_estoque = estoque_ant - qtd
            custo_medio = custo_ant

        estoques[prod] = novo_estoque
        custos[prod] = custo_medio

        df.at[i,"Quantidade em Estoque"] = novo_estoque
        df.at[i,"Custo Médio"] = custo_medio
        df.at[i,"Valor Total em Estoque"] = novo_estoque * custo_medio

    return df

@st.dialog("Excluir Movimentação")
def excluir_movimentacao():

    try:
        df = pd.read_excel(arquivo)
        df["Data"] = pd.to_datetime(df["Data"])
    except:
        st.warning("Nenhuma movimentação encontrada.")
        return

    df["MovID"] = df.index

    movimento = st.selectbox(
        "Selecione a movimentação",
        df["MovID"],
        format_func=lambda x: f'{df.loc[x,"Data"].strftime("%d/%m/%Y")} - {df.loc[x,"Produto"]} - {df.loc[x,"Tipo"]} ({df.loc[x,"Quantidade"]})'
    )

    if st.button("🗑 Excluir"):

        df = df.drop(index=movimento)

        df = recalcular_estoque(df)

        df.to_excel(arquivo, index=False)

        st.success("Movimentação excluída!")
        st.rerun()
    
@st.dialog("Excluir Despesa")
def excluir_despesa():

    try:
        despesas = pd.read_excel("Despesas.xlsx")
        despesas["Data"] = pd.to_datetime(despesas["Data"])
    except:
        st.warning("Nenhuma despesa encontrada.")
        return

    despesas["DespID"] = despesas.index

    despesa = st.selectbox(
        "Selecione a despesa",
        despesas["DespID"],
        format_func=lambda x: f'{despesas.loc[x,"Data"].strftime("%d/%m/%Y")} - {despesas.loc[x,"Descrição"]} - {moeda_br(despesas.loc[x,"Valor"])}'
    )

    if st.button("🗑 Excluir Despesa"):

        despesas = despesas.drop(index=despesa)

        despesas.to_excel("Despesas.xlsx", index=False)

        st.success("Despesa excluída!")
        st.rerun()

# SIDEBAR
st.sidebar.write(f"👤 Usuário: {name}")
if st.sidebar.button("➕ Cadastrar Movimentação"):
    cadastrar_produto()

if st.sidebar.button("🗑 Excluir Movimentação"):
    excluir_movimentacao()

if st.sidebar.button("💸 Cadastrar Despesa"):
    cadastrar_despesa()

if st.sidebar.button("🗑 Excluir Despesa"):
    excluir_despesa()

if authentication_status:
    authenticator.logout("🚪 Logout", "sidebar")

# SISTEMA ESTOQUE
df = pd.read_excel(arquivo)
df["Data"] = pd.to_datetime(df["Data"])

st.title("📦 Sistema Estoque")

# CALCULOS
vendas = df[df["Tipo"] == "Venda"]
compras = df[df["Tipo"] == "Compra"]

qtd_vendida = vendas["Quantidade"].sum()
qtd_comprada = compras["Quantidade"].sum()

valor_vendido = vendas["Valor Total"].sum()
valor_comprado = compras["Valor Total"].sum()

# DESPESAS
try:
    despesas = pd.read_excel("Despesas.xlsx")
    despesas["Data"] = pd.to_datetime(despesas["Data"])
except:
    despesas = pd.DataFrame(columns=["Data","Descrição","Categoria","Valor"])

despesa_total = despesas["Valor"].sum() if not despesas.empty else 0


# RESUMO
st.subheader("Resumo")

col1, col2, col3 = st.columns(3)

col1.metric("🛒 Quantidade Vendida", f"{qtd_vendida:,}")
col2.metric("📦 Quantidade Comprada", f"{qtd_comprada:,}")

estoque_atual = df.sort_values("Data").groupby("Produto").tail(1)
quantidade_estoque = estoque_atual["Quantidade em Estoque"].sum()

col3.metric("📦 Quantidade em Estoque", f"{int(quantidade_estoque):,}")


col4, col5, col6 = st.columns(3)

col4.metric("💰 Valor Vendido", moeda_br(valor_vendido))
col5.metric("🛍️ Valor Comprado", moeda_br(valor_comprado))

lucro = valor_vendido - valor_comprado

delta_cor = "normal"
if lucro < 0:
    delta_cor = "inverse"

col6.metric(
    label="📈 Lucro Bruto",
    value=f"R$ {lucro:,.2f}",
    delta=f"R$ {lucro:,.2f}",
    delta_color=delta_cor
)

# TABELA
df_view = df.copy()
colunas_visiveis_estoque = [
    "Data", "Produto", "Tipo", "Cliente/Fornecedor", 
    "Quantidade", "Preço Unitário", "Valor Total", 
    "Quantidade em Estoque", "Custo Médio", "Valor Total em Estoque"
]
df_view["Preço Unitário"] = df_view["Preço Unitário"].apply(moeda_br)
df_view["Valor Total"] = df_view["Valor Total"].apply(moeda_br)
df_view["Custo Médio"] = df_view["Custo Médio"].apply(moeda_br)
df_view["Valor Total em Estoque"] = df_view["Valor Total em Estoque"].apply(moeda_br)
st.dataframe(
    df_view[colunas_visiveis_estoque], 
    column_config={
        "Data": st.column_config.DateColumn("Data", format="DD/MM/YYYY")
    },
    use_container_width=True,
    hide_index=True
)

# VENDAS POR PRODUTO
st.title("💸 Resumo de Vendas por Produto")

vendas_produto = vendas.groupby("Produto").agg({
    "Quantidade": "sum",
    "Valor Total": "sum"
}).reset_index()

col1, col2 = st.columns(2)

with col1:
    st.subheader("Quantidade por Produto")

    grafico_qtd = alt.Chart(vendas_produto).mark_bar(
        color="red",
        size=30
    ).encode(
        x="Produto:N",
        y="Quantidade:Q"
    ).properties(height=500)

    st.altair_chart(grafico_qtd, use_container_width=True)

with col2:
    st.subheader("Valor por Produto")

    grafico_valor = alt.Chart(vendas_produto).mark_bar(
        color="red",
        size=30
    ).encode(
        x="Produto:N",
        y="Valor Total:Q"
    ).properties(height=500)

    st.altair_chart(grafico_valor, use_container_width=True)

# DESPESAS
st.title("📉 Despesas")
despesas_view = despesas.copy()

if not despesas_view.empty:
    colunas_visiveis_desp = ["Data", "Descrição", "Fornecedor", "Categoria", "Valor"]
    despesas_view["Data"] = despesas_view["Data"].dt.strftime("%d/%m/%Y")
    despesas_view["Valor"] = despesas_view["Valor"].apply(moeda_br)
    st.dataframe(despesas_view[colunas_visiveis_desp], use_container_width=True, hide_index=True)
else:
    st.info("Nenhuma despesa cadastrada.")

# GRAFICO DESPESAS
st.subheader("Despesas por Categoria")

if not despesas.empty:
    despesas_categoria = despesas.groupby("Categoria")["Valor"].sum().reset_index()

    grafico_despesa = alt.Chart(despesas_categoria).mark_bar(
        color="orange",
        size=30
    ).encode(
        x="Categoria:N",
        y="Valor:Q"
    ).properties(height=400)

    st.altair_chart(grafico_despesa, use_container_width=True)

# DRE
st.title("📑 DRE")

receita_total = valor_vendido
cmv = valor_comprado

lucro_bruto = receita_total - cmv
lucro_liquido = lucro_bruto - despesa_total

col1, col2, col3, col4 = st.columns(4)

col1.metric("💰 Receita", moeda_br(receita_total))
col2.metric("📦 Custo das Compras", moeda_br(cmv))
col3.metric("📉 Despesas", moeda_br(despesa_total))
col4.metric("📈 Lucro Líquido", moeda_br(lucro_liquido))

# CONTAS A PAGAR E RECEBER
st.divider()
st.title("🏦 Fluxo de Caixa")

receber_total = vendas["Valor Total"].sum()

pagar_compras = compras["Valor Total"].sum()
pagar_despesas = despesas["Valor"].sum()
pagar_total = pagar_compras + pagar_despesas

c1, c2, c3 = st.columns(3)

with c1:
    st.subheader("💰 A Receber")
    st.write("*(Total de Vendas)*")
    st.metric("Total Clientes", moeda_br(receber_total))

with c2:
    st.subheader("💸 A Pagar")
    st.write("*(Compras + Despesas)*")
    st.metric("Total Fornecedores", moeda_br(pagar_total), delta_color="inverse")

with c3:
    st.subheader("⚖️ Saldo Previsto")
    st.write("*(Receber - Pagar)*")
    saldo_final = receber_total - pagar_total
    st.metric("Resultado", moeda_br(saldo_final), delta=f"{saldo_final:,.2f}")

tab1, tab2 = st.tabs(["A Receber", "A Pagar"])

with tab1:
    st.markdown("### Vendas por Data (Clientes)")
    if not vendas.empty:
        df_receber = vendas[["Data", "Produto", "Quantidade", "Valor Total"]].copy()
        df_receber["Data"] = df_receber["Data"].dt.strftime("%d/%m/%Y")
        df_receber["Valor Total"] = df_receber["Valor Total"].apply(moeda_br)
        st.dataframe(df_receber, use_container_width=True, hide_index=True)
    else:
        st.info("Sem vendas registradas.")

with tab2:
    st.markdown("### Compras e Despesas (Fornecedores)")

    if not compras.empty or not despesas.empty:
        comp_fin = compras[["Data", "Produto", "Valor Total"]].rename(columns={"Produto": "Descrição", "Valor Total": "Valor"})
        desp_fin = despesas[["Data", "Descrição", "Valor"]]
       
        df_pagar_total = pd.concat([comp_fin, desp_fin], ignore_index=True)
        df_pagar_total = df_pagar_total.sort_values("Data", ascending=False)
        
        df_pagar_total["Data"] = pd.to_datetime(df_pagar_total["Data"]).dt.strftime("%d/%m/%Y")
        df_pagar_total["Valor"] = df_pagar_total["Valor"].apply(moeda_br)
        
        st.dataframe(df_pagar_total, use_container_width=True, hide_index=True)
    else:
        st.info("Sem contas a pagar registradas.")