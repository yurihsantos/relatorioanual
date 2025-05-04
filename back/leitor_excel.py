# Importa bibliotecas necessárias 
import pandas as pd # type: ignore
import locale
from utils import ordinal

# Define a localização de formatação.
locale.setlocale(locale.LC_ALL, "pt_BR.UTF-8")

local_origem = ""
local_destino = ""
livro_num = ""
df = pd.DataFrame()
db = pd.DataFrame()

def inputsalvar():
    # Recebe o nome da pasta de trabalho
    global local_origem, local_destino
    local_origem = "C:\\Users\\Yurih Santos\\Documents\\dados.xlsx"
    local_destino = "C:\\Users\\Yurih Santos\\Documents\\saída.xlsx"
    return 

def data_primario():
    # Cria e formata o dataframe primário
    global df
    df = pd.read_excel(local_origem)

    # Ajusta colunas conforme a localização, limpa espaços duplos e espaços no começo e no final de células
    df["Data"] = df["Data"].dt.strftime("%d/%m/%Y")
    df["Hora"] = pd.to_datetime(df["Hora"], format="%H:%M").dt.strftime("%H:%M")
    df = df.apply(lambda col: col.str.strip() if col.dtype == "object" else col)

    #Cria novas colunas na base de dados original para facilitar as funções
    df["TotalLeilao"] = df["Código"].map(df.groupby("Código")["Valor de Venda"].sum())
    df["SomaComissão"] = df["Valor de Venda"] + df["Comissão"]
    df["Base"] = (
        df["Praça"].apply(ordinal) + 
        " leilão " + 
        df["Modalidade"].str.lower() + 
        " realizado em " + 
        df["Local"] + 
        " na data " + 
        df["Data"] + 
        " às " + 
        df["Hora"] + 
        "h. Leilão " + 
        df["Tipo"].str.lower() + " (" + 
        df["Processo"].apply(lambda x: "sem processo vinculado" if x == "0000000-00.0000.0.00.0000" else f"vinculado ao processo nº {x}") + 
        ") cujo comitente é " + 
        df["Comitente"] + "."
    )

    df["Descrição_Livro"] = (
        df["Descrição"] + ". Valor mínimo de " +
        df["Valor Mínimo"].apply(lambda x: locale.currency(x, grouping=True)) + ". " +
        df.apply(lambda row: "Lote não arrematado." if row["Valor de Venda"] == 0 
        else "Arrematado por " + row["Arrematante"] + " no valor de " + 
        locale.currency(row["Valor de Venda"], grouping=True) + ".", axis=1)
    )
    return

# Cria um novo dataframe e cria as colunas respectivas de cada livro.
def L3():
    global livro_num, db
    livro_num = "3"
    db["LLL"] = ("Lote " +  df["Lote"].astype(str))
    db["TL3"] = df["Base"].copy()
    db["DL3"] = df["Descrição_Livro"].copy()
    return

def L4():
    global livro_num, db
    livro_num = "4"
    db["LLL"] = ("Lote " +  df["Lote"].astype(str))
    db["TL4"] = (
        df["Base"] + " Leiloeiro nomeado em " + df["Nomeação"] + 
        df["Prestação"].apply(lambda x: " e não houve prestação de contas." if x == "XXXXX" else f" e prestou contas em {x}.")
    )
    db["DL4"] = df["Descrição"].copy()
    return

def L5():
    global livro_num, db, df
    livro_num = "5"
    db["LLL"] = ("Lote " +  df["Lote"].astype(str))
    db["TL5"] = (
        df["Base"] + 
        df["TotalLeilao"].apply(lambda x: f" Produto bruto do leilão equivalente a {locale.currency(x, grouping=True)}" + "." if x != 0 else "")
    )
    db["DL5"] = (
        df["Descrição_Livro"].copy() + 
        df["SomaComissão"].apply(lambda x: f" A arrematação e a comissão totalizam {locale.currency(x, grouping=True)}" + "." if x != 0 else "")
    )
    return

def limpardb():
    global df, db, dt, local_destino, livro_num

    # Limpa os itens redundantes das linhas cujo lote é diferente de 1
    dt = db.groupby("TL" + livro_num)[["LLL", "DL"+ livro_num]].apply(lambda x: "; ".join(x["LLL"] + " - " + x["DL" + livro_num]) + ".").reset_index(name="DL" + livro_num)
    dt["DL" + livro_num] = dt["DL"+ livro_num].str.replace(".;", ";").str.replace("..", ".")

    # Exporta a planilha
    dt.to_excel(local_destino, index=False)
    return

inputsalvar()
data_primario()
L3()
limpardb()