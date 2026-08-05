"""
Baixa a lista completa (todas as linhas) do painel Power BI de CADIFA da Anvisa.

Como funciona:
- Replica a mesma requisição POST que o navegador faz para /public/reports/querydata.
- Aumenta o parâmetro de paginação ("Window Count") para trazer mais de 500 linhas.
- Decodifica o formato compactado (DSR) que o Power BI usa nas respostas
  (valores repetidos entre linhas não são reenviados; strings de alta repetição
  vêm como índice de um dicionário; strings únicas vêm por extenso).

Requisitos:
    pip install requests pandas openpyxl

Uso:
    python baixar_cadifa.py
"""

import json
import uuid
import requests
import pandas as pd

# ---------------------------------------------------------------------------
# Configuração fixa do relatório (extraída da requisição capturada no DevTools)
# ---------------------------------------------------------------------------
RESOURCE_KEY = "940d6cea-7507-417a-97d1-7ea436d3a113"
TENANT_ID = "b67af23f-c3f3-4d35-80c7-b7085f5edd81"
QUERY_URL = "https://wabi-brazil-south-api.analysis.windows.net/public/reports/querydata?synchronous=true"

# Quantas linhas pedir por vez. Comece com um valor alto; se o servidor
# recusar ou truncar, reduza (ex.: 5000, depois 2000).
PAGE_SIZE = 10000


def montar_payload(window_count: int) -> dict:
    """Monta o corpo da requisição, igual ao capturado no navegador,
    mas com o Window.Count ajustado para trazer mais linhas."""
    query = {
        "Version": 2,
        "From": [
            {"Name": "t", "Entity": "TA_DADOS_CADIFA", "Type": 0},
            {"Name": "t1", "Entity": "TA_HISTORICO_PETICAO", "Type": 0},
        ],
        "Select": [
            {"Column": {"Expression": {"SourceRef": {"Source": "t"}}, "Property": "NO_RAZAO_SOCIAL_MAISC"},
             "Name": "TA_DADOS_CADIFA.NO_RAZAO_SOCIAL_MAISC", "NativeReferenceName": "Razão Social da Empresa"},
            {"Column": {"Expression": {"SourceRef": {"Source": "t"}}, "Property": "NO_INSUMO_MINUSC"},
             "Name": "TA_DADOS_CADIFA.NO_INSUMO_MINUSC", "NativeReferenceName": "Nome do Insumo"},
            {"Column": {"Expression": {"SourceRef": {"Source": "t"}}, "Property": "NU_PROCESSO"},
             "Name": "TA_DADOS_CADIFA.NU_PROCESSO", "NativeReferenceName": "Nº CADIFA"},
            {"Column": {"Expression": {"SourceRef": {"Source": "t"}}, "Property": "DS_APRESENTACAO_PRODUTO"},
             "Name": "TA_DADOS_CADIFA.DS_APRESENTACAO_PRODUTO", "NativeReferenceName": "Revisão"},
            {"Aggregation": {"Expression": {"Column": {"Expression": {"SourceRef": {"Source": "t1"}},
                                                        "Property": "DT_FIM_SITUACAO"}}, "Function": 4},
             "Name": "TA_HISTORICO_PETICAO.DT_FIM_SITUACAO", "NativeReferenceName": "Data da Última Situação"},
        ],
        "Where": [
            {"Condition": {"In": {"Expressions": [{"Column": {"Expression": {"SourceRef": {"Source": "t"}},
                                                               "Property": "CO_ASSUNTO"}}],
                                   "Values": [[{"Literal": {"Value": "'11637'"}}]]}}},
            {"Condition": {"In": {"Expressions": [{"Column": {"Expression": {"SourceRef": {"Source": "t"}},
                                                               "Property": "DS_SITUACAO_APRESENTACAO"}}],
                                   "Values": [[{"Literal": {"Value": "'Deferida'"}}]]}}},
            {"Condition": {"In": {"Expressions": [{"Column": {"Expression": {"SourceRef": {"Source": "t"}},
                                                               "Property": "DS_ASSUNTO"}}],
                                   "Values": [
                                       [{"Literal": {"Value": "'CADIFA - Solicitação de CADIFA Associada | Associated CADIFA Application'"}}],
                                       [{"Literal": {"Value": "'CADIFA - Solicitação de CADIFA Não Associada | Standalone CADIFA Application'"}}],
                                   ]}}},
        ],
        "OrderBy": [{"Direction": 2, "Expression": {"Column": {"Expression": {"SourceRef": {"Source": "t"}},
                                                                "Property": "DS_APRESENTACAO_PRODUTO"}}}],
    }

    command = {
        "SemanticQueryDataShapeCommand": {
            "Query": query,
            "Binding": {
                "Primary": {"Groupings": [{"Projections": [0, 1, 2, 3, 4]}]},
                "DataReduction": {"DataVolume": 3, "Primary": {"Window": {"Count": window_count}}},
                "Version": 1,
            },
            "ExecutionMetricsKind": 1,
        }
    }

    body = {
        "version": "1.0.0",
        "queries": [{
            "Query": {"Commands": [command]},
            "QueryId": "",
            "ApplicationContext": {
                "DatasetId": "0dd556db-ae50-4cf0-957e-566ccee995ac",
                "Sources": [{"ReportId": "1dd397d6-880e-418c-8808-22138d08da99",
                             "VisualId": "7c9bea16044e5ad9ddcc"}],
            },
        }],
        "cancelQueries": [],
        "modelId": 8373560,
    }
    return body


def buscar_dados(window_count: int) -> dict:
    headers = {
        "Content-Type": "application/json;charset=UTF-8",
        "Accept": "application/json",
        "X-PowerBI-ResourceKey": RESOURCE_KEY,
        "ActivityId": str(uuid.uuid4()),
        "RequestId": str(uuid.uuid4()),
        "Referer": "https://app.powerbi.com/",
        "Origin": "https://app.powerbi.com",
        "User-Agent": "Mozilla/5.0",
    }
    resp = requests.post(QUERY_URL, headers=headers, data=json.dumps(montar_payload(window_count)))
    resp.raise_for_status()
    return resp.json()


def decodificar_dsr(resposta: dict) -> pd.DataFrame:
    """Decodifica o formato compactado DSR do Power BI em uma lista de linhas."""
    result = resposta["results"][0]["result"]["data"]
    dsr = result["dsr"]
    value_dicts = dsr.get("DS", [{}])[0].get("ValueDicts", {})
    rows_raw = dsr["DS"][0]["PH"][0]["DM0"]

    colunas = None  # lista de dicts {N, DN(opcional)}
    linhas_decodificadas = []
    linha_anterior = None

    for linha in rows_raw:
        if "S" in linha:
            colunas = linha["S"]  # define ordem e dicionários das colunas

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

        # Resolve índices de dicionário -> valor real (quando aplicável)
        linha_final = []
        for i, col in enumerate(colunas):
            val = nova_linha[i]
            dn = col.get("DN")
            if dn and isinstance(val, int):
                val = value_dicts[dn][val]
            linha_final.append(val)

        linhas_decodificadas.append(linha_final)
        linha_anterior = nova_linha

    nomes_colunas = ["Razão Social", "Insumo (IFA)", "Nº CADIFA", "Revisão", "Data Última Situação (epoch ms)"]
    df = pd.DataFrame(linhas_decodificadas, columns=nomes_colunas)

    # Converte a coluna de data (epoch em milissegundos) para data legível
    df["Data Última Situação"] = pd.to_datetime(df["Data Última Situação (epoch ms)"], unit="ms").dt.strftime("%d/%m/%Y")
    df = df.drop(columns=["Data Última Situação (epoch ms)"])

    return df


def main():
    print(f"Buscando até {PAGE_SIZE} linhas...")
    resposta = buscar_dados(PAGE_SIZE)
    df = decodificar_dsr(resposta)
    print(f"Total de linhas obtidas: {len(df)}")

    if len(df) == PAGE_SIZE:
        print("ATENÇÃO: o número de linhas retornado é igual ao solicitado.")
        print("Isso pode indicar que ainda existem mais linhas (limite atingido).")
        print("Aumente PAGE_SIZE no topo do script e rode novamente.")

    saida = "cadifa_completo.xlsx"
    df.to_excel(saida, index=False)
    print(f"Arquivo salvo: {saida}")


if __name__ == "__main__":
    main()
