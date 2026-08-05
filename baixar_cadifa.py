"""
Baixa a lista completa de CADIFA da Anvisa capturando DINAMICAMENTE a
consulta real que o navegador faz ao abrir o painel (via Playwright),
em vez de hardcodar nomes de coluna/entidade.

Por que isso importa:
    A versao anterior deste script tinha os nomes das colunas
    (NO_RAZAO_SOCIAL, NO_INSUMO, CO_ASSUNTO, etc.) fixos no codigo.
    Quando a Anvisa mudou o modelo de dados do painel, esses nomes
    deixaram de existir e o script parou de funcionar
    (erro "CouldNotResolveSemanticQueryDefinition").

    Esta versao evita isso: ela abre o painel de verdade num navegador
    headless, captura a consulta que o PROPRIO painel envia (sempre com
    os nomes atuais, corretos), e so ajusta o tamanho da pagina antes
    de repetir essa mesma consulta via requests. Assim, o script se
    adapta automaticamente a mudancas de nome de coluna. Ele só quebra
    se a Anvisa mudar a ESTRUTURA VISUAL do painel (ex.: remover a
    tabela inteira), o que é bem mais raro.

Requisitos:
    pip install playwright requests pandas openpyxl
    playwright install chromium --with-deps

Uso:
    python baixar_cadifa.py
"""

import json
import uuid
import requests
import pandas as pd
from playwright.sync_api import sync_playwright

REPORT_URL = (
    "https://app.powerbi.com/view?r=eyJrIjoiOTQwZDZjZWEtNzUwNy00MTdhLTk3ZDEtN2VhNDM2ZDNhMTEzIiwidCI6ImI2N2FmMjNmLWMzZjMtNGQzNS04MGM3LWI3MDg1ZjVlZGQ4MSJ9"
)

# Quantas linhas pedir. Comece alto; reduza se o servidor recusar.
PAGE_SIZE = 10000


def capturar_consulta_real() -> dict:
    """Abre o painel num navegador headless e captura a primeira
    requisicao real de querydata que ele mesmo dispara -- com os
    nomes de coluna ATUAIS do modelo de dados da Anvisa."""
    capturado = {}

    def handle_request(request):
        if "querydata" in request.url and request.method == "POST":
            if "payload" not in capturado:
                capturado["url"] = request.url
                capturado["payload"] = request.post_data
                capturado["headers"] = dict(request.headers)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.on("request", handle_request)

        page.goto(REPORT_URL, wait_until="load", timeout=60000)

        tentativas = 0
        while "payload" not in capturado and tentativas < 20:
            page.wait_for_timeout(1000)
            tentativas += 1

        browser.close()

    if "payload" not in capturado:
        raise RuntimeError(
            "Nao foi possivel capturar a requisicao querydata do painel. "
            "O layout/estrutura visual do painel pode ter mudado -- "
            "seria necessario inspecionar manualmente via DevTools."
        )

    return capturado


def extrair_nomes_amigaveis(payload_str: str) -> list:
    """Le do payload capturado os nomes de exibicao (NativeReferenceName)
    de cada coluna selecionada, na ordem em que aparecem."""
    payload = json.loads(payload_str)
    try:
        query = payload["queries"][0]["Query"]["Commands"][0]["SemanticQueryDataShapeCommand"]["Query"]
        selects = query["Select"]
    except (KeyError, IndexError, TypeError):
        return None

    nomes = []
    for item in selects:
        nome = item.get("NativeReferenceName") or item.get("Name") or "coluna"
        nomes.append(nome)
    return nomes


def aumentar_paginacao(payload_str: str, window_count: int) -> str:
    """Troca o valor de Window.Count no payload capturado, mantendo o
    resto da consulta (com os nomes de coluna atuais) intacto."""
    payload = json.loads(payload_str)

    def ajustar(obj):
        if isinstance(obj, dict):
            if isinstance(obj.get("Window"), dict) and "Count" in obj["Window"]:
                obj["Window"]["Count"] = window_count
            for v in obj.values():
                ajustar(v)
        elif isinstance(obj, list):
            for v in obj:
                ajustar(v)

    ajustar(payload)
    return json.dumps(payload)


def buscar_dados(url: str, payload_str: str, headers_originais: dict) -> dict:
    headers = {
        "Content-Type": "application/json;charset=UTF-8",
        "Accept": "application/json",
        "ActivityId": str(uuid.uuid4()),
        "RequestId": str(uuid.uuid4()),
        "Referer": "https://app.powerbi.com/",
        "Origin": "https://app.powerbi.com",
        "User-Agent": "Mozilla/5.0",
    }
    for chave in ("x-powerbi-resourcekey", "X-PowerBI-ResourceKey"):
        if chave in headers_originais:
            headers["X-PowerBI-ResourceKey"] = headers_originais[chave]
            break

    resp = requests.post(url, headers=headers, data=payload_str)
    resp.raise_for_status()
    return resp.json()


def decodificar_dsr(resposta: dict, nomes_amigaveis: list = None) -> pd.DataFrame:
    """Decodifica o formato compactado DSR do Power BI. Nao depende dos
    nomes tecnicos das colunas -- so da ordem/estrutura da resposta."""
    result = resposta["results"][0]["result"]["data"]
    dsr = result["dsr"]
    value_dicts = dsr.get("DS", [{}])[0].get("ValueDicts", {})
    rows_raw = dsr["DS"][0]["PH"][0]["DM0"]

    colunas = None
    linhas_decodificadas = []
    linha_anterior = None

    for linha in rows_raw:
        if "S" in linha:
            colunas = linha["S"]

        valores_atuais = linha.get("C", [])
        bitmask_repete = linha.get("R", 0)

        n_col = len(colunas)
        nova_linha = [None] * n_col
        idx_valor = 0
        for i in range(n_col):
            repete = bool(bitmask_repete & (1 << i))
            if repete and linha_anterior is not None:
                nova_linha[i] = linha_anterior[i]
            else:
                nova_linha[i] = valores_atuais[idx_valor]
                idx_valor += 1

        linha_final = []
        for i, col in enumerate(colunas):
            val = nova_linha[i]
            dn = col.get("DN")
            if dn and isinstance(val, int):
                val = value_dicts[dn][val]
            linha_final.append(val)

        linhas_decodificadas.append(linha_final)
        linha_anterior = nova_linha

    if nomes_amigaveis and len(nomes_amigaveis) == len(colunas):
        nomes_finais = nomes_amigaveis
    else:
        nomes_finais = [col.get("N", f"coluna_{i}") for i, col in enumerate(colunas)]

    df = pd.DataFrame(linhas_decodificadas, columns=nomes_finais)

    # Converte colunas de data (formato "dd/MM/yyyy" no schema) de epoch ms
    for i, col in enumerate(colunas):
        fmt = col.get("Format", "")
        if isinstance(fmt, str) and "d" in fmt.lower() and "y" in fmt.lower():
            nome_col = nomes_finais[i]
            try:
                df[nome_col] = pd.to_datetime(df[nome_col], unit="ms").dt.strftime("%d/%m/%Y")
            except Exception:
                pass  # se nao for epoch, deixa como esta

    return df


def main():
    print("Abrindo o painel da Anvisa para capturar a consulta atual...")
    capturado = capturar_consulta_real()
    print("Consulta capturada com sucesso.")

    nomes_amigaveis = extrair_nomes_amigaveis(capturado["payload"])
    payload_ajustado = aumentar_paginacao(capturado["payload"], PAGE_SIZE)

    print(f"Buscando até {PAGE_SIZE} linhas...")
    resposta = buscar_dados(capturado["url"], payload_ajustado, capturado["headers"])
    df = decodificar_dsr(resposta, nomes_amigaveis)
    print(f"Total de linhas obtidas: {len(df)}")
    print(f"Colunas obtidas: {list(df.columns)}")

    if len(df) == PAGE_SIZE:
        print("ATENÇÃO: quantidade retornada igual ao solicitado.")
        print("Pode haver mais linhas -- aumente PAGE_SIZE e rode novamente.")

    saida = "cadifa_completo.xlsx"
    df.to_excel(saida, index=False)
    print(f"Arquivo salvo: {saida}")


if __name__ == "__main__":
    main()
