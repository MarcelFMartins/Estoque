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
            "Data","Produto","Tipo","Cliente/Fornecedor","Quantidade","Preço Unitário",
            "Valor Total","Quantidade em Estoque","Custo Médio","Valor Total em Estoque"
        ])

        produtos_cadastrados = ["Morango","Tomate"]

    produto_mov = st.selectbox("Produto", produtos_cadastrados)
    tipo_mov = st.selectbox("Tipo", ["Compra","Venda"])

    # CLIENTE OU FORNECEDOR
    if tipo_mov == "Compra":

        try:
            fornecedores = pd.read_excel("Fornecedores.xlsx")
            contato_valor = st.selectbox("Fornecedor", fornecedores["Nome"])
        except:
            st.warning("Cadastre um fornecedor primeiro.")
            st.stop()

    else:

        try:
            clientes = pd.read_excel("Clientes.xlsx")
            contato_valor = st.selectbox("Cliente", clientes["Nome"])
        except:
            st.warning("Cadastre um cliente primeiro.")
            st.stop()

    quantidade_mov = st.number_input("Quantidade", min_value=1)
    preco_mov = st.number_input("Preço Unitário", min_value=0.0, format="%.2f")

    forma_pagamento = st.selectbox("Forma de Pagamento", ["À Vista", "A Prazo"])

    if forma_pagamento == "A Prazo":

        parcelas = st.number_input("Número de Parcelas", min_value=2)

        data_recebimento = st.date_input("Data do Primeiro Vencimento")

    else:

        parcelas = 1
        data_recebimento = data_mov

    if st.button("Salvar Cadastro"):

        try:

            valor_total = quantidade_mov * preco_mov

            nova_linha = pd.DataFrame([{
                "Data": pd.to_datetime(data_mov),
                "Produto": produto_mov,
                "Tipo": tipo_mov,
                "Cliente/Fornecedor": contato_valor,
                "Quantidade": quantidade_mov,
                "Preço Unitário": preco_mov,
                "Valor Total": valor_total,
                "Forma Pagamento": forma_pagamento,
                "Parcelas": parcelas,
                "Data Recebimento": data_recebimento,
                "Status": "Recebido" if forma_pagamento == "À Vista" else "Pendente"
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

            st.success("Movimentação cadastrada com sucesso!")
            st.rerun()

        except Exception as e:
            st.error(f"Erro ao salvar: {e}")

# CADASTRAR DESPESA
@st.dialog("Cadastrar Despesa")
def cadastrar_despesa():

    data_desp = st.date_input("Data", value=date.today())
    descricao = st.text_input("Descrição")

    categoria = st.selectbox(
        "Categoria",
        ["Combustível", "Alimentação", "Hotel", "Despesas com Funcionários", "Outros"]
    )

    valor = st.number_input("Valor", min_value=0.0)

    # FORNECEDOR
    try:
        fornecedores = pd.read_excel("Fornecedores.xlsx")
        fornecedor = st.selectbox("Fornecedor", fornecedores["Nome"])
    except:
        st.warning("Cadastre um fornecedor primeiro.")
        st.stop()

    forma_pagamento = st.selectbox("Forma de Pagamento", ["À Vista","A Prazo"])

    if forma_pagamento == "A Prazo":

        parcelas = st.number_input("Parcelas", min_value=1, value=1)

        data_recebimento = st.date_input("Data Pagamento")

        status = "Pendente"

    else:

        parcelas = 1
        data_recebimento = data_desp
        status = "Pago"

    if st.button("Salvar Despesa"):

        nova_despesa = pd.DataFrame({
            "Data": [data_desp],
            "Descrição": [descricao],
            "Cliente/Fornecedor": [fornecedor],
            "Categoria": [categoria],
            "Valor": [valor],
            "Forma Pagamento": [forma_pagamento],
            "Parcelas": [parcelas],
            "Data Recebimento": [data_recebimento],
            "Status": [status]
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

@st.dialog("Cadastrar Cliente")
def cadastrar_cliente():

    nome = st.text_input("Nome do Cliente")

    if st.button("Salvar Cliente"):

        novo = pd.DataFrame({"Nome":[nome]})

        try:
            df = pd.read_excel("Clientes.xlsx")
            df = pd.concat([df,novo],ignore_index=True)
        except:
            df = novo

        df.to_excel("Clientes.xlsx",index=False)

        st.success("Cliente cadastrado")
        st.rerun()

@st.dialog("Cadastrar Fornecedor")
def cadastrar_fornecedor():

    nome = st.text_input("Nome do Fornecedor")

    if st.button("Salvar Fornecedor"):

        novo = pd.DataFrame({"Nome":[nome]})

        try:
            df = pd.read_excel("Fornecedores.xlsx")
            df = pd.concat([df,novo],ignore_index=True)
        except:
            df = novo

        df.to_excel("Fornecedores.xlsx",index=False)

        st.success("Fornecedor cadastrado")
        st.rerun()

@st.dialog("Excluir Cliente ou Fornecedor")
def excluir_contato():

    tipo = st.selectbox("Tipo", ["Cliente", "Fornecedor"])

    if tipo == "Cliente":

        try:
            df = pd.read_excel("Clientes.xlsx")
        except:
            st.warning("Nenhum cliente cadastrado.")
            return

    else:

        try:
            df = pd.read_excel("Fornecedores.xlsx")
        except:
            st.warning("Nenhum fornecedor cadastrado.")
            return

    contato = st.selectbox("Selecione para excluir", df["Nome"])

    if st.button("🗑 Excluir"):

        df = df[df["Nome"] != contato]

        if tipo == "Cliente":
            df.to_excel("Clientes.xlsx", index=False)
        else:
            df.to_excel("Fornecedores.xlsx", index=False)

        st.success(f"{tipo} excluído com sucesso!")
        st.rerun()

# SIDEBAR



st.sidebar.subheader("Movimentações e Despesas")
if st.sidebar.button("➕ Cadastrar Movimentação"):
    cadastrar_produto()

if st.sidebar.button("🗑 Excluir Movimentação"):
    excluir_movimentacao()

if st.sidebar.button("💸 Cadastrar Despesa"):
    cadastrar_despesa()

if st.sidebar.button("🗑 Excluir Despesa"):
    excluir_despesa()

st.sidebar.divider()
st.sidebar.subheader("Clientes e Fornecedores")

if st.sidebar.button("👤 Cadastrar Cliente"):
    cadastrar_cliente()

if st.sidebar.button("🏭 Cadastrar Fornecedor"):
    cadastrar_fornecedor()

st.sidebar.button("🗑 Excluir Cliente/Fornecedor", on_click=excluir_contato)

# SISTEMA ESTOQUE
df = pd.read_excel(arquivo)

colunas_novas = ["Forma Pagamento","Parcelas","Data Recebimento","Status"]

for col in colunas_novas:
    if col not in df.columns:
        df[col] = ""

df["Data"] = pd.to_datetime(df["Data"])

st.sidebar.divider()
st.sidebar.subheader("📅 Período")

data_inicio = st.sidebar.date_input(
    "Data Inicial",
    value=df["Data"].min()
)

data_fim = st.sidebar.date_input(
    "Data Final",
    value=df["Data"].max()
)

df = df[
    (df["Data"] >= pd.to_datetime(data_inicio)) &
    (df["Data"] <= pd.to_datetime(data_fim))
]

st.sidebar.divider()
st.sidebar.write(f"👤 Usuário: {name}")
if authentication_status:
    authenticator.logout("🚪 Logout", "sidebar")

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

df_ordenado = df.sort_values("Data")
valor_estoque = df_ordenado.iloc[-1]["Valor Total em Estoque"]

# RESUMO
st.subheader("📊 Resumo")

produtos = df["Produto"].unique()

for produto in produtos:

    df_prod = df[df["Produto"] == produto]

    vendas_prod = df_prod[df_prod["Tipo"] == "Venda"]
    compras_prod = df_prod[df_prod["Tipo"] == "Compra"]

    qtd_vendida = vendas_prod["Quantidade"].sum()
    qtd_comprada = compras_prod["Quantidade"].sum()

    estoque_atual = df_prod.sort_values("Data").iloc[-1]["Quantidade em Estoque"]

    valor_vendido_prod = vendas_prod["Valor Total"].sum()

    st.markdown(f"### {produto}")

    col1, col2, col3, col4 = st.columns(4)

    col1.metric("🛒 Vendido", f"{int(qtd_vendida):,}")
    col2.metric("📦 Comprado", f"{int(qtd_comprada):,}")
    col3.metric("📦 Estoque", f"{int(estoque_atual):,}")
    col4.metric("💰 Valor Vendido", moeda_br(valor_vendido_prod))

    st.divider()


col5, col6, col7 = st.columns(3)

col5.metric("💰 Valor Vendido", moeda_br(valor_vendido))
col6.metric("🛍️ Valor Comprado", moeda_br(valor_comprado))
col7.metric("💵 Valor em Estoque", moeda_br(valor_estoque))



# TABELA
df_view = df.copy()
colunas_visiveis_estoque = [
    "Data", "Produto", "Tipo", "Cliente/Fornecedor", 
    "Quantidade", "Preço Unitário", "Valor Total", 
    "Quantidade em Estoque", "Custo Médio", "Valor Total em Estoque",
    "Forma Pagamento", "Parcelas", "Data Recebimento", "Status"
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
    colunas_visiveis_desp = ["Data", "Descrição", "Categoria", "Valor"]
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
cmv = valor_comprado - valor_estoque

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

receber = df[(df["Tipo"] == "Venda") & (df["Status"] == "Pendente")]
recebido = df[(df["Tipo"] == "Venda") & (df["Status"] == "Recebido")]

pagar = df[(df["Tipo"] == "Compra") & (df["Status"] == "Pendente")]
pago = df[(df["Tipo"] == "Compra") & (df["Status"] == "Pago")]

despesas_pendentes = despesas[despesas["Status"] == "Pendente"]
despesas_pagas = despesas[despesas["Status"] == "Pago"]

receber_total = receber["Valor Total"].sum()

pagar_compras = pagar["Valor Total"].sum()
pagar_despesas = despesas_pendentes["Valor"].sum()

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

    st.markdown("### Clientes")

    receber = df[(df["Tipo"] == "Venda") & (df["Status"] == "Pendente")]
    recebido = df[(df["Tipo"] == "Venda") & (df["Status"] == "Recebido")]

    pagar = df[(df["Tipo"] == "Compra") & (df["Status"] == "Pendente")]
    pago = df[(df["Tipo"] == "Compra") & (df["Status"] == "Pago")]

    col1, col2 = st.columns(2)

    with col1:

        st.subheader("🟡 A Receber")

        if not receber.empty:

            df_receber = receber[[
                "Data",
                "Cliente/Fornecedor",
                "Produto",
                "Quantidade",
                "Valor Total",
                "Forma Pagamento"
            ]].copy()

            df_receber["Data"] = df_receber["Data"].dt.strftime("%d/%m/%Y")
            df_receber["Valor Total"] = df_receber["Valor Total"].apply(moeda_br)

            st.dataframe(df_receber, use_container_width=True, hide_index=True)

        else:
            st.info("Nenhuma conta a receber.")

    with col2:

        st.subheader("🟢 Recebido")

        if not recebido.empty:

            df_recebido = recebido[[
                "Data",
                "Cliente/Fornecedor",
                "Produto",
                "Quantidade",
                "Valor Total",
                "Forma Pagamento"
            ]].copy()

            df_recebido["Data"] = df_recebido["Data"].dt.strftime("%d/%m/%Y")
            df_recebido["Valor Total"] = df_recebido["Valor Total"].apply(moeda_br)

            st.dataframe(df_recebido, use_container_width=True, hide_index=True)

        else:
            st.info("Nenhum recebimento registrado.")

with tab2:

    st.markdown("### Fornecedores")

    pagar = df[(df["Tipo"] == "Compra") & (df["Status"] == "Pendente")]
    pago = df[(df["Tipo"] == "Compra") & (df["Status"] == "Pago")]

    col1, col2 = st.columns(2)

    with col1:

        st.subheader("🔴 A Pagar")

        if not pagar.empty:
            df_pagar = pagar[[
                "Data",
                "Cliente/Fornecedor",
                "Produto",
                "Quantidade",
                "Valor Total",
                "Forma Pagamento"
            ]].copy()

            df_despesas_pend = despesas_pendentes[[
                "Data",
                "Cliente/Fornecedor",
                "Descrição",
                "Valor",
                "Forma Pagamento"
            ]].copy()

            df_despesas_pend = df_despesas_pend.rename(columns={
                "Descrição": "Produto",
                "Valor": "Valor Total"
            })

            df_total_pagar = pd.concat([df_pagar, df_despesas_pend])

            df_total_pagar["Data"] = df_total_pagar["Data"].dt.strftime("%d/%m/%Y")
            df_total_pagar["Valor Total"] = df_total_pagar["Valor Total"].apply(moeda_br)

            st.dataframe(df_total_pagar, use_container_width=True, hide_index=True)
        
        else:
            st.info("Nenhuma conta a pagar.")

    with col2:

        st.subheader("🟢 Pago")
        if not pago.empty:
            df_pago = pago[[
                "Data",
                "Cliente/Fornecedor",
                "Produto",
                "Quantidade",
                "Valor Total",
                "Forma Pagamento"
            ]].copy()

            df_despesas_pag = despesas_pagas[[
                "Data",
                "Cliente/Fornecedor",
                "Descrição",
                "Valor",
                "Forma Pagamento"
            ]].copy()

            df_despesas_pag = df_despesas_pag.rename(columns={
                "Descrição": "Produto",
                "Valor": "Valor Total"
            })

            df_total_pago = pd.concat([df_pago, df_despesas_pag])

            df_total_pago["Data"] = df_total_pago["Data"].dt.strftime("%d/%m/%Y")
            df_total_pago["Valor Total"] = df_total_pago["Valor Total"].apply(moeda_br)

            st.dataframe(df_total_pago, use_container_width=True, hide_index=True)

        else:
            st.info("Nenhum pagamento registrado.")