"""
Baixa a lista de CADIFAs com status "Deferida" do Painel Cadifa da Anvisa
(Power BI publico), decodificando o formato compactado da resposta.

Esta versao usa a consulta EXATA capturada do painel em 06/08/2026,
apos a Anvisa ter reestruturado o painel (nova tela com filtros +
botao "Pesquisar"). Os nomes de coluna atuais sao:
    NO_RAZAO_SOCIAL_MAISC, NO_INSUMO_MINUSC, NU_PROCESSO,
    DS_APRESENTACAO_PRODUTO, DT_FIM_SITUACAO (agregacao MAX)
Filtro atual: DS_SITUACAO_APRESENTACAO = 'Deferida' (unico filtro --
o painel novo nao usa mais o filtro de CO_ASSUNTO/DS_ASSUNTO da
versao antiga).

Se a Anvisa mudar o esquema de novo, o sintoma sera um erro
"CouldNotResolveSemanticQueryDefinition". Nesse caso, é necessario
capturar a consulta atualizada de novo via DevTools (F12 -> Network
-> filtrar "querydata" -> Payload) e atualizar a funcao
montar_payload() abaixo.

Requisitos:
    pip install requests pandas openpyxl

Uso:
    python baixar_cadifa.py
"""

import json
import uuid
import requests
import pandas as pd

RESOURCE_KEY = "940d6cea-7507-417a-97d1-7ea436d3a113"
TENANT_ID = "b67af23f-c3f3-4d35-80c7-b7085f5edd81"
QUERY_URL = "https://wabi-brazil-south-api.analysis.windows.net/public/reports/querydata?synchronous=true"

# Quantas linhas pedir. Hoje (06/08/2026) ha 578 deferidas -- deixamos
# bem acima disso para folga.
PAGE_SIZE = 10000


def montar_payload(window_count: int) -> dict:
    query = {
        "Version": 2,
        "From": [
            {"Name": "t", "Entity": "TA_DADOS_CADIFA", "Type": 0},
            {"Name": "t1", "Entity": "TA_HISTORICO_PETICAO", "Type": 0},
        ],
        "Select": [
            {"Column": {"Expression": {"SourceRef": {"Source": "t"}}, "Property": "NO_RAZAO_SOCIAL_MAISC"},
             "Name": "TA_DADOS_CADIFA.NO_RAZAO_SOCIAL_MAISC", "NativeReferenceName": "Razão Social da Empresa1"},
            {"Column": {"Expression": {"SourceRef": {"Source": "t"}}, "Property": "NO_INSUMO_MINUSC"},
             "Name": "TA_DADOS_CADIFA.NO_INSUMO_MINUSC", "NativeReferenceName": "Nome do Insumo"},
            {"Column": {"Expression": {"SourceRef": {"Source": "t"}}, "Property": "NU_PROCESSO"},
             "Name": "TA_DADOS_CADIFA.NU_PROCESSO", "NativeReferenceName": "Nº Cadifa"},
            {"Column": {"Expression": {"SourceRef": {"Source": "t"}}, "Property": "DS_APRESENTACAO_PRODUTO"},
             "Name": "TA_DADOS_CADIFA.DS_APRESENTACAO_PRODUTO", "NativeReferenceName": "Revisão"},
            {"Aggregation": {"Expression": {"Column": {"Expression": {"SourceRef": {"Source": "t1"}},
                                                        "Property": "DT_FIM_SITUACAO"}}, "Function": 4},
             "Name": "TA_HISTORICO_PETICAO.DT_FIM_SITUACAO", "NativeReferenceName": "Data da Última Situação"},
        ],
        "Where": [
            {"Condition": {"In": {"Expressions": [{"Column": {"Expression": {"SourceRef": {"Source": "t"}},
                                                               "Property": "DS_SITUACAO_APRESENTACAO"}}],
                                   "Values": [[{"Literal": {"Value": "'Deferida'"}}]]}}},
        ],
        "OrderBy": [{"Direction": 2, "Expression": {"Aggregation": {
            "Expression": {"Column": {"Expression": {"SourceRef": {"Source": "t1"}},
                                       "Property": "DT_FIM_SITUACAO"}},
            "Function": 4}}}],
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

    return {
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
    corpo = resp.json()

    resultados = corpo.get("results", [])
    if resultados:
        erro = resultados[0].get("result", {}).get("data", {}).get("error")
        if erro:
            raise RuntimeError(
                f"A Anvisa rejeitou a consulta (provavel mudanca de esquema de "
                f"colunas). Detalhe: {erro}. E necessario capturar a consulta "
                f"atualizada via DevTools e atualizar montar_payload()."
            )
    return corpo


def decodificar_dsr(resposta: dict) -> pd.DataFrame:
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

    nomes_colunas = ["Razão Social", "Insumo (IFA)", "Nº CADIFA", "Revisão", "Data Última Situação (epoch ms)"]
    df = pd.DataFrame(linhas_decodificadas, columns=nomes_colunas)

    df["Data Última Situação"] = pd.to_datetime(df["Data Última Situação (epoch ms)"], unit="ms").dt.strftime("%d/%m/%Y")
    df = df.drop(columns=["Data Última Situação (epoch ms)"])

    return df


def main():
    print(f"Buscando até {PAGE_SIZE} linhas...")
    resposta = buscar_dados(PAGE_SIZE)
    df = decodificar_dsr(resposta)
    print(f"Total de linhas obtidas: {len(df)}")

    if len(df) == PAGE_SIZE:
        print("ATENÇÃO: quantidade retornada igual ao solicitado.")
        print("Pode haver mais linhas -- aumente PAGE_SIZE e rode novamente.")

    saida = "cadifa_completo.xlsx"
    df.to_excel(saida, index=False)
    print(f"Arquivo salvo: {saida}")


if __name__ == "__main__":
    main()
