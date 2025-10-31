import os
from datetime import datetime
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from urllib.parse import urlparse, parse_qs

# Bibliotecas para Google Sheets
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from dotenv import load_dotenv   # 🔹 NOVO: para carregar variáveis do .env

# 🔹 NOVO: carregar variáveis do arquivo .env logo no início
load_dotenv()

CLIENTES_CSV = "clientes.csv"
HISTORICO_CSV = "historico_versoes.csv"
HISTORICO_XLSX = "historico_versoes.xlsx"
GOOGLE_SHEET_NAME = "historico_versoes"  # Nome da planilha no Google Drive
# CREDENCIAIS_JSON = "credenciais.json"  # ❌ REMOVIDO: não usamos mais caminho fixo

def extrair_versao_com_selenium(url: str) -> str | None:
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")

    driver = webdriver.Chrome(options=options)
    try:
        driver.get(url)
        final_url = driver.current_url
        parsed = urlparse(final_url)
        params = parse_qs(parsed.query)
        if "v" in params and len(params["v"]) > 0:
            return f"v={params['v'][0]}"
        return None
    finally:
        driver.quit()

def carregar_historico():
    if os.path.exists(HISTORICO_CSV):
        return pd.read_csv(HISTORICO_CSV, encoding="cp1252")
    else:
        return pd.DataFrame(columns=["Cliente", "URL", "Versão Atual", "Versão Anterior", "Data da pesquisa"])

def salvar_historico(df):
    # Salva CSV
    df.to_csv(HISTORICO_CSV, index=False, encoding="cp1252")

    # Salva Excel com formatação
    with pd.ExcelWriter(HISTORICO_XLSX, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="Histórico", index=False)

        workbook  = writer.book
        worksheet = writer.sheets["Histórico"]

        header_format = workbook.add_format({
            "bold": True,
            "text_wrap": True,
            "valign": "top",
            "fg_color": "#D7E4BC",
            "border": 1
        })

        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
            worksheet.set_column(col_num, col_num, 30)

        worksheet.autofilter(0, 0, len(df), len(df.columns)-1)

    # Também envia para Google Sheets
    salvar_google_sheets(df)

def salvar_google_sheets(df):
    """
    Atualiza a planilha no Google Sheets com os dados do DataFrame.
    """
    scope = ["https://spreadsheets.google.com/feeds",
             "https://www.googleapis.com/auth/drive"]

    # 🔹 ALTERADO: agora pegamos o caminho do JSON da variável de ambiente
    cred_path = os.getenv("GOOGLE_APPLICATION_CREDENTIALS")
    creds = ServiceAccountCredentials.from_json_keyfile_name(cred_path, scope)
    client = gspread.authorize(creds)

    spreadsheet = client.open(GOOGLE_SHEET_NAME)
    worksheet = spreadsheet.sheet1

    worksheet.clear()
    worksheet.update([df.columns.values.tolist()] + df.values.tolist())

def atualizar_registro(df, cliente, url, nova_versao):
    agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if cliente in df["Cliente"].values:
        idx = df.index[df["Cliente"] == cliente][0]
        versao_atual = df.at[idx, "Versão Atual"]

        if nova_versao and nova_versao != versao_atual:
            df.at[idx, "Versão Anterior"] = versao_atual
            df.at[idx, "Versão Atual"] = nova_versao
            df.at[idx, "Data da pesquisa"] = agora
        else:
            df.at[idx, "Data da pesquisa"] = agora
    else:
        df.loc[len(df)] = [cliente, url, nova_versao, nova_versao, agora]

    return df

def processar():
    df_clientes = pd.read_csv(CLIENTES_CSV)
    df_hist = carregar_historico()

    for _, row in df_clientes.iterrows():
        cliente = row["Cliente"]
        url = row["URL"]
        versao = extrair_versao_com_selenium(url) or "Versão não encontrada"
        df_hist = atualizar_registro(df_hist, cliente, url, versao)

    salvar_historico(df_hist)

def main():
    processar()
    print("Coleta concluída. Dados salvos em CSV, Excel e Google Sheets.")

if __name__ == "__main__":
    main()