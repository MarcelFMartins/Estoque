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
        },
        "christian": {
            "name": "Christian",
            "password": "B2uaa5cl*"
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

def numero_br(valor):
    return f"{valor:,.0f}".replace(",", ".")

# CADASTRAR MOVIMENTAÇÃO
@st.dialog("Cadastrar Movimentação")
def cadastrar_produto():

    data_mov = st.date_input("Data", value=date.today())

    try:
        df = pd.read_excel(arquivo)

        df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
        df["Produto"] = df["Produto"].astype(str)
        df["Tipo"] = df["Tipo"].astype(str)

        try:
            produtos = pd.read_excel("Produtos.xlsx")
            produtos_cadastrados = produtos["Produto"].tolist()
        except:
            produtos_cadastrados = []

    except:
        df = pd.DataFrame(columns=[
            "Data","Produto","Tipo","Cliente/Fornecedor","Quantidade","Preço Unitário",
            "Valor Total","Quantidade em Estoque","Custo Médio","Valor Total em Estoque"
        ])

        produtos_cadastrados = ["Morango","Tomate"]

    produto_mov = st.selectbox("Produto", produtos_cadastrados)

    estoque_atual = 0

    if not df.empty:
        df_prod = df[df["Produto"] == produto_mov]
        if not df_prod.empty:
            estoque_atual = df_prod.sort_values("Data").iloc[-1]["Quantidade em Estoque"]

    st.info(f"Estoque atual: {numero_br(estoque_atual)}")

    tipo_mov = st.selectbox(
        "Tipo",
        ["Compra","Venda","Perda"]
    )


    # CLIENTE OU FORNECEDOR
    if tipo_mov == "Compra" or "Perda":

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

    if tipo_mov == "Venda" and quantidade_mov > estoque_atual:
        st.error("❌ Estoque insuficiente para essa venda.")
        st.stop()

    preco_mov = st.number_input("Preço Unitário", min_value=0.0, format="%.2f")

    if tipo_mov == "Perda":

        forma_pagamento = "Perda"
        parcelas = 1
        parcelas_pagas = 1
        data_recebimento = data_mov

    else:
        forma_pagamento = st.selectbox(
            "Forma de Pagamento",
            ["À Vista", "A Prazo"]
        )

    if forma_pagamento == "A Prazo":

        parcelas = st.number_input(
            "Número de Parcelas",
            min_value=2,
            step=1
        )

        parcelas_pagas = st.number_input(
            "Parcelas já pagas",
            min_value=0,
            max_value=int(parcelas),
            step=1
        )

        data_recebimento = st.date_input(
            "Data do Primeiro Vencimento"
        )

    else:

        parcelas = 1
        parcelas_pagas = 1
        data_recebimento = data_mov

    if st.button("Salvar Cadastro"):

        try:

            valor_total = quantidade_mov * preco_mov

            valor_parcela = valor_total / parcelas
            valor_pago = valor_parcela * parcelas_pagas
            valor_restante = valor_total - valor_pago

            if valor_restante == 0:
                status = "Recebido" if tipo_mov == "Venda" else "Pago"
            else:
                status = "Pendente"

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
                "Parcelas Pagas": parcelas_pagas,

                "Valor Pago": valor_pago,
                "Valor Restante": valor_restante,

                "Data Recebimento": data_recebimento,
                "Status": status
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

                elif tipo == "Venda":

                    novo_estoque = estoque_ant - qtd
                    custo_medio = custo_ant

                elif tipo == "Perda":

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

        elif tipo == "Venda":

            novo_estoque = estoque_ant - qtd
            custo_medio = custo_ant

        elif tipo == "Perda":

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

@st.dialog("Excluir Cliente")
def excluir_cliente():

    try:
        df = pd.read_excel("Clientes.xlsx")
    except:
        st.warning("Nenhum cliente cadastrado.")
        return

    cliente = st.selectbox("Selecione o cliente", df["Nome"])

    if st.button("🗑 Excluir Cliente"):

        df = df[df["Nome"] != cliente]

        df.to_excel("Clientes.xlsx", index=False)

        st.success("Cliente excluído com sucesso!")
        st.rerun()

@st.dialog("Excluir Fornecedor")
def excluir_fornecedor():

    try:
        df = pd.read_excel("Fornecedores.xlsx")
    except:
        st.warning("Nenhum fornecedor cadastrado.")
        return

    fornecedor = st.selectbox("Selecione o fornecedor", df["Nome"])

    if st.button("🗑 Excluir Fornecedor"):

        df = df[df["Nome"] != fornecedor]

        df.to_excel("Fornecedores.xlsx", index=False)

        st.success("Fornecedor excluído com sucesso!")
        st.rerun()

@st.dialog("Cadastrar Produto")
def cadastrar_produto_novo():

    nome = st.text_input("Nome do Produto")
    categoria = st.text_input("Categoria")
    estoque_min = st.number_input("Estoque Mínimo", min_value=0)
    unidade = st.selectbox("Unidade", ["KG","UN","CX","L"])

    if st.button("Salvar Produto"):

        novo = pd.DataFrame({
            "Produto":[nome],
            "Categoria":[categoria],
            "Estoque Mínimo":[estoque_min],
            "Unidade":[unidade]
        })

        try:
            df = pd.read_excel("Produtos.xlsx")
            df = pd.concat([df,novo],ignore_index=True)
        except:
            df = novo

        df.to_excel("Produtos.xlsx",index=False)

        st.success("Produto cadastrado!")
        st.rerun()

@st.dialog("Excluir Produto")
def excluir_produto():

    try:
        df = pd.read_excel("Produtos.xlsx")
    except:
        st.warning("Nenhum produto cadastrado.")
        return

    produto = st.selectbox("Selecione o produto", df["Produto"])

    if st.button("🗑 Excluir Produto"):

        df = df[df["Produto"] != produto]

        df.to_excel("Produtos.xlsx", index=False)

        st.success("Produto excluído!")
        st.rerun()

@st.dialog("Editar Produto")
def editar_produto():

    try:
        df = pd.read_excel("Produtos.xlsx")
    except:
        st.warning("Nenhum produto cadastrado.")
        return

    produto = st.selectbox("Selecione o produto", df["Produto"])

    if produto:

        linha = df[df["Produto"] == produto].index[0]

        nome = st.text_input("Nome", value=df.loc[linha,"Produto"])
        categoria = st.text_input("Categoria", value=df.loc[linha,"Categoria"])
        estoque_min = st.number_input(
            "Estoque Mínimo",
            value=int(df.loc[linha,"Estoque Mínimo"])
        )
        unidade = st.selectbox(
            "Unidade",
            ["KG","UN","CX","L"],
            index=["KG","UN","CX","L"].index(df.loc[linha,"Unidade"])
        )

        if st.button("Salvar Alterações"):

            df.loc[linha,"Produto"] = nome
            df.loc[linha,"Categoria"] = categoria
            df.loc[linha,"Estoque Mínimo"] = estoque_min
            df.loc[linha,"Unidade"] = unidade

            df.to_excel("Produtos.xlsx", index=False)

            st.success("Produto atualizado!")
            st.rerun()

@st.dialog("Editar Movimentação")
def editar_movimentacao():

    try:
        df = pd.read_excel(arquivo)
        df["Data"] = pd.to_datetime(df["Data"])
    except:
        st.warning("Nenhuma movimentação encontrada.")
        return

    df["MovID"] = df.index

    try:
        df_clientes = pd.read_excel("Clientes.xlsx")
        lista_clientes = df_clientes["Nome"].dropna().unique().tolist()
    except:
        lista_clientes = []

    try:
        df_fornec = pd.read_excel("Fornecedores.xlsx")
        lista_fornec = df_fornec["Nome"].dropna().unique().tolist()
    except:
        lista_fornec = []

    mov = st.selectbox(
        "Selecione a movimentação",
        df["MovID"],
        format_func=lambda x: f'{df.loc[x,"Data"].strftime("%d/%m/%Y")} - {df.loc[x,"Produto"]} - {df.loc[x,"Tipo"]} ({df.loc[x,"Quantidade"]})'
    )

    linha = df.loc[mov]

    data_mov = st.date_input("Data", value=linha["Data"])

    produto = st.text_input("Produto", value=linha["Produto"])

    tipos = ["Compra","Venda","Perda"]

    tipo = st.selectbox(
        "Tipo",
        tipos,
        index=tipos.index(linha["Tipo"]) if linha["Tipo"] in tipos else 0
    )

    if tipo == "Compra":
    
        opcoes = lista_fornec

        if linha["Cliente/Fornecedor"] in opcoes:
            idx = opcoes.index(linha["Cliente/Fornecedor"])
        else:
            idx = 0 if opcoes else None

        contato = st.selectbox(
            "Fornecedor",
            opcoes,
            index=idx if opcoes else None,
            key="edit_fornec"
        )


    elif tipo == "Venda":

        opcoes = lista_clientes

        if linha["Cliente/Fornecedor"] in opcoes:
            idx = opcoes.index(linha["Cliente/Fornecedor"])
        else:
            idx = 0 if opcoes else None

        contato = st.selectbox(
            "Cliente",
            opcoes,
            index=idx if opcoes else None,
            key="edit_cli"
        )

    elif tipo == "Perda":

        opcoes = lista_fornec

        if linha["Cliente/Fornecedor"] in opcoes:
            idx = opcoes.index(linha["Cliente/Fornecedor"])
        else:
            idx = 0 if opcoes else None

        contato = st.selectbox(
            "Fornecedor",
            opcoes,
            index=idx if opcoes else None,
            key="edit_fornec"
        )

    quantidade = st.number_input(
        "Quantidade",
        value=int(linha["Quantidade"])
    )

    preco = st.number_input(
        "Preço Unitário",
        value=float(linha["Preço Unitário"])
    )

    forma_pagamento = st.selectbox(
        "Forma de Pagamento",
        ["À Vista", "A Prazo"],
        index=0 if linha["Forma Pagamento"] == "À Vista" else 1
    )

    if forma_pagamento == "A Prazo":

        parcelas = st.number_input(
            "Parcelas",
            min_value=1,
            value=int(linha.get("Parcelas", 1))
        )

        parcelas_pagas = st.number_input(
            "Parcelas Pagas",
            min_value=0,
            max_value=int(parcelas),
            value=int(linha.get("Parcelas Pagas", 0))
        )

        data_recebimento = st.date_input(
            "Data Recebimento",
            value=pd.to_datetime(
                linha.get("Data Recebimento", date.today())
            )
        )

    else:

        parcelas = 1
        parcelas_pagas = 1
        data_recebimento = data_mov

    if st.button("Salvar Alterações"):

        preco_unit = preco  # ← garante que existe

        valor_total = quantidade * preco_unit

        valor_parcela = valor_total / parcelas
        valor_pago = valor_parcela * parcelas_pagas
        valor_restante = valor_total - valor_pago

        if valor_restante == 0:
            status = "Recebido" if tipo == "Venda" else "Pago"
        else:
            status = "Pendente"

        df.loc[mov,"Data"] = data_mov
        df.loc[mov,"Produto"] = produto
        df.loc[mov,"Tipo"] = tipo
        df.loc[mov,"Cliente/Fornecedor"] = contato
        df.loc[mov,"Quantidade"] = quantidade
        df.loc[mov,"Preço Unitário"] = preco_unit
        df.loc[mov,"Valor Total"] = valor_total

        df.loc[mov,"Forma Pagamento"] = forma_pagamento
        df.loc[mov,"Parcelas"] = parcelas
        df.loc[mov,"Parcelas Pagas"] = parcelas_pagas
        df.loc[mov,"Valor Pago"] = valor_pago
        df.loc[mov,"Valor Restante"] = valor_restante
        df.loc[mov,"Data Recebimento"] = data_recebimento
        df.loc[mov,"Status"] = status

        df = recalcular_estoque(df)

        df.to_excel(arquivo,index=False)

        st.success("Movimentação atualizada!")
        st.rerun()

@st.dialog("Editar Despesa")
def editar_despesa():

    try:
        df = pd.read_excel("Despesas.xlsx")
        df["Data"] = pd.to_datetime(df["Data"])
    except:
        st.warning("Nenhuma despesa encontrada.")
        return
    
    try:
        df_fornec = pd.read_excel("Fornecedores.xlsx")
        lista_fornec = df_fornec["Nome"].dropna().unique().tolist()
    except:
        lista_fornec = []

    df["DespID"] = df.index

    desp = st.selectbox(
        "Selecione a despesa",
        df["DespID"],
        format_func=lambda x: f'{df.loc[x,"Data"].strftime("%d/%m/%Y")} - {df.loc[x,"Descrição"]}'
    )

    linha = df.loc[desp]

    data = st.date_input("Data", value=linha["Data"])
    descricao = st.text_input(
        "Descrição",
        value=linha.get("Descrição","")
    )

    categoria = st.text_input(
        "Categoria",
        value=linha.get("Categoria","")
    )


    # ===== FORNECEDOR SELECT =====

    fornecedor_atual = linha.get("Fornecedor","")

    if fornecedor_atual in lista_fornec:
        idx = lista_fornec.index(fornecedor_atual)
    else:
        idx = 0 if lista_fornec else None

    fornecedor = st.selectbox(
        "Fornecedor",
        lista_fornec,
        index=idx if lista_fornec else None,
        key="edit_desp_fornec"
    )


    valor = st.number_input(
        "Valor",
        value=float(linha["Valor"])
    )

    if st.button("Salvar Alterações"):

        df.loc[desp,"Descrição"] = descricao
        df.loc[desp,"Categoria"] = categoria
        df.loc[desp,"Fornecedor"] = fornecedor
        df.loc[desp,"Valor"] = valor
        df.to_excel("Despesas.xlsx",index=False)

        st.success("Despesa atualizada!")
        st.rerun()

@st.dialog("Editar Cliente")
def editar_cliente():

    try:
        df = pd.read_excel("Clientes.xlsx")
    except:
        st.warning("Nenhum cliente cadastrado.")
        return

    cliente = st.selectbox("Cliente", df["Nome"])

    novo_nome = st.text_input("Novo Nome", value=cliente)

    if st.button("Salvar Alteração"):

        df.loc[df["Nome"]==cliente,"Nome"] = novo_nome

        df.to_excel("Clientes.xlsx",index=False)

        st.success("Cliente atualizado!")
        st.rerun()

@st.dialog("Editar Fornecedor")
def editar_fornecedor():

    try:
        df = pd.read_excel("Fornecedores.xlsx")
    except:
        st.warning("Nenhum fornecedor cadastrado.")
        return

    fornecedor = st.selectbox("Fornecedor", df["Nome"])

    novo_nome = st.text_input("Novo Nome", value=fornecedor)

    if st.button("Salvar Alteração"):

        df.loc[df["Nome"]==fornecedor,"Nome"] = novo_nome

        df.to_excel("Fornecedores.xlsx",index=False)

        st.success("Fornecedor atualizado!")
        st.rerun()

# SIDEBAR



st.sidebar.subheader("Movimentações")
if st.sidebar.button("➕ Cadastrar Movimentação"):
    cadastrar_produto()

st.sidebar.button("✏️ Editar Movimentação", on_click=editar_movimentacao)

if st.sidebar.button("🗑 Excluir Movimentação"):
    excluir_movimentacao()

st.sidebar.divider()
st.sidebar.subheader("Despesas")
if st.sidebar.button("💸 Cadastrar Despesa"):
    cadastrar_despesa()

st.sidebar.button("✏️ Editar Despesa", on_click=editar_despesa)

if st.sidebar.button("🗑 Excluir Despesa"):
    excluir_despesa()

st.sidebar.divider()
st.sidebar.subheader("Clientes")

if st.sidebar.button("👤 Cadastrar Cliente"):
    cadastrar_cliente()

st.sidebar.button("✏️ Editar Cliente", on_click=editar_cliente)

st.sidebar.button("🗑 Excluir Cliente", on_click=excluir_cliente)

st.sidebar.divider()
st.sidebar.subheader("Fornecedores")
if st.sidebar.button("🏭 Cadastrar Fornecedor"):
    cadastrar_fornecedor()

st.sidebar.button("✏️ Editar Fornecedor", on_click=editar_fornecedor)

st.sidebar.button("🗑 Excluir Fornecedor", on_click=excluir_fornecedor)

st.sidebar.divider()
st.sidebar.subheader("Produtos")

if st.sidebar.button("➕ Cadastrar Produto"):
    cadastrar_produto_novo()

st.sidebar.button("✏️ Editar Produto", on_click=editar_produto)

if st.sidebar.button("🗑 Excluir Produto"):
    excluir_produto()

# SISTEMA ESTOQUE
try:
    df = pd.read_excel(arquivo)
except:
    df = pd.DataFrame()

colunas_novas = [
    "Forma Pagamento",
    "Parcelas",
    "Parcelas Pagas",
    "Valor Pago",
    "Valor Restante",
    "Data Recebimento",
    "Status"
]

for col in colunas_novas:
    if col not in df.columns:
        df[col] = ""

df["Data"] = pd.to_datetime(df["Data"])

st.sidebar.divider()
st.sidebar.subheader("📅 Período")

df["Data"] = pd.to_datetime(df["Data"], errors="coerce")

if df["Data"].dropna().empty:
    data_inicio_padrao = date.today()
    data_fim_padrao = date.today()
else:
    data_inicio_padrao = df["Data"].min()
    data_fim_padrao = df["Data"].max()

data_inicio = st.sidebar.date_input(
    "Data Inicial",
    value=data_inicio_padrao
)

data_fim = st.sidebar.date_input(
    "Data Final",
    value=data_fim_padrao
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

try:
    produtos = pd.read_excel("Produtos.xlsx")

    ultimos = df.sort_values("Data").groupby("Produto").last().reset_index()

    alerta = ultimos.merge(produtos, on="Produto", how="left")

    estoque_baixo = alerta[
        alerta["Quantidade em Estoque"] <= alerta["Estoque Mínimo"]
    ]

    if not estoque_baixo.empty:

        st.warning("⚠ Produtos com estoque baixo")

        for _, row in estoque_baixo.iterrows():
            st.write(
                f"🔸 {row['Produto']} — Estoque: {numero_br(row['Quantidade em Estoque'])} (Mínimo: {numero_br(row['Estoque Mínimo'])})"
            )

except:
    pass

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

if not df.empty:
    df_ordenado = df.sort_values("Data")
    valor_estoque = df_ordenado.iloc[-1]["Valor Total em Estoque"]
else:
    valor_estoque = 0

st.divider()

st.subheader("Análise por Produto")

produto_selecionado = st.selectbox(
    "Selecione o produto",
    produtos,
    index=None,
    placeholder="Escolha um produto",
    key="analise_produto"
)

if produto_selecionado:

    df_prod = df[df["Produto"] == produto_selecionado]

    vendas_prod = df_prod[df_prod["Tipo"] == "Venda"]
    compras_prod = df_prod[df_prod["Tipo"] == "Compra"]
    perdas_prod = df_prod[df_prod["Tipo"] == "Perda"]

    qtd_vendida = vendas_prod["Quantidade"].sum()
    valor_vendido_prod = vendas_prod["Valor Total"].sum()

    qtd_comprada = compras_prod["Quantidade"].sum()
    valor_comprado_prod = compras_prod["Valor Total"].sum()

    qtd_perdida = perdas_prod["Quantidade"].sum()
    valor_perdido_prod = perdas_prod["Valor Total"].sum()

    if not df_prod.empty:
        estoque_qtd = df_prod.sort_values("Data").iloc[-1]["Quantidade em Estoque"]
        estoque_valor = df_prod.sort_values("Data").iloc[-1]["Valor Total em Estoque"]
    else:
        estoque_qtd = 0
        estoque_valor = 0

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.metric("🛒 Quantidade Vendida", numero_br(qtd_vendida))
        st.metric("💰 Valor Vendido", moeda_br(valor_vendido_prod))

    with col2:
        st.metric("📥 Quantidade Comprada", numero_br(qtd_comprada))
        st.metric("🛍️ Valor Comprado", moeda_br(valor_comprado_prod))

    with col3:
        st.metric("📦 Estoque Atual", numero_br(estoque_qtd))
        st.metric("💵 Valor em Estoque", moeda_br(estoque_valor))

    with col4:
        st.metric("⚠️ Quantidade Perdida", numero_br(qtd_perdida))
        st.metric("💸 Valor Perdido", moeda_br(valor_perdido_prod))

st.divider()
st.subheader("Resumo Geral")
col5, col6, col7, col8 = st.columns(4)

col5.metric("💰 Valor Total Vendido", moeda_br(valor_vendido))
col6.metric("🛍️ Valor Total Comprado", moeda_br(valor_comprado))

ultimos_produtos = df.sort_values("Data").groupby("Produto").last()
valore_em_estoque = ultimos_produtos["Valor Total em Estoque"].sum()

total_estoque = ultimos_produtos["Quantidade em Estoque"].sum()

col7.metric("💵 Valor Total em Estoque", moeda_br(valore_em_estoque))
col8.metric("📦 Quantidade Total em Estoque", numero_br(total_estoque))

st.divider()
st.subheader("Perdas Geral")

perdas = df[df["Tipo"] == "Perda"]

valor_perdido_total = perdas["Valor Total"].sum()
qtd_perdida_total = perdas["Quantidade"].sum()

colp1, colp2 = st.columns(2)

with colp2:
    st.metric(
        "⚠️ Quantidade Total Perdida",
        numero_br(qtd_perdida_total)
    )

with colp1:
    st.metric(
        "💸 Valor Total Perdido",
        moeda_br(valor_perdido_total)
    )

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

st.divider()

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

st.divider()

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

st.divider()

# DRE
st.title("📑 DRE")

receita_total = valor_vendido
perdas = df[df["Tipo"] == "Perda"]["Valor Total"].sum()

cmv = valor_comprado - valore_em_estoque

lucro_bruto = receita_total - cmv
lucro_liquido = lucro_bruto - despesa_total - perdas

col1, col2, col3, col4, col5 = st.columns(5)

col1.metric("💰 Receita", moeda_br(receita_total))
col2.metric("📦 Custo das Compras", moeda_br(cmv))
col3.metric("📉 Despesas", moeda_br(despesa_total))
col4.metric("⚠️ Perdas", moeda_br(valor_perdido_total))
col5.metric("📈 Lucro Líquido", moeda_br(lucro_liquido))

# CONTAS A PAGAR E RECEBER
st.divider()
st.title("🏦 Fluxo de Caixa")

receber = df[(df["Tipo"] == "Venda") & (df["Status"] == "Pendente")]
recebido = df[(df["Tipo"] == "Venda") & (df["Status"] == "Recebido")]

pagar = df[(df["Tipo"] == "Compra") & (df["Status"] == "Pendente")]
pago = df[(df["Tipo"] == "Compra") & (df["Status"] == "Pago")]

despesas_pendentes = despesas[despesas["Status"] == "Pendente"]
despesas_pagas = despesas[despesas["Status"] == "Pago"]

receber_total = receber["Valor Restante"].sum()

pagar_compras = pagar["Valor Restante"].sum()
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
    st.metric("Resultado", moeda_br(saldo_final), delta=moeda_br(saldo_final))

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