"""
Baixa a lista completa (todas as linhas) do painel Power BI de CADIFA da Anvisa.

DIFERENÇA em relação à versão anterior:
- Em vez de fixar manualmente o ResourceKey, DatasetId, ReportId, VisualId e o
  corpo inteiro da query (que quebram sempre que a Anvisa republica o
  relatório), este script abre a URL pública num navegador headless
  (Playwright) e INTERCEPTA a requisição real que o próprio Power BI faz
  para /public/reports/querydata.
- O único ajuste feito na requisição interceptada é aumentar o
  "Window.Count" (paginação), para trazer mais de 500 linhas.
- O decodificador do formato DSR também foi tornado mais genérico: em vez de
  nomes de coluna fixos, ele lê os nomes a partir do próprio payload da
  requisição (NativeReferenceName), então se a Anvisa adicionar/remover uma
  coluna do relatório, o script não quebra silenciosamente - ele reflete a
  mudança nas colunas do Excel de saída.

Requisitos:
    pip install playwright pandas openpyxl
    playwright install chromium

Uso:
    python baixar_cadifa.py
"""

import json
import re
import sys

import pandas as pd
from playwright.sync_api import sync_playwright

# ---------------------------------------------------------------------------
# Configuração
# ---------------------------------------------------------------------------

REPORT_URL = (
    "https://app.powerbi.com/view?r="
    "eyJrIjoiOTQwZDZjZWEtNzUwNy00MTdhLTk3ZDEtN2VhNDM2ZDNhMTEzIiwidCI6ImI2N2FmMjNmLWMzZjMtNGQzNS04MGM3LWI3MDg1ZjVlZGQ4MSJ9"
)

# Trecho único que identifica a requisição do visual certo (o que consulta
# TA_DADOS_CADIFA), para diferenciar de outras chamadas querydata que a
# página possa fazer para outros visuais/filtros.
MARCADOR_ENTIDADE = "TA_DADOS_CADIFA"

# Quantas linhas pedir. Comece alto; se a resposta vier truncada em exatamente
# esse número, aumente e rode de novo.
PAGE_SIZE = 10000

TIMEOUT_MS = 45_000


def _ajustar_window_count(payload_bytes: bytes, novo_count: int) -> bytes:
    """Recebe o corpo (JSON) da requisição real capturada no navegador e
    substitui apenas o Window.Count, mantendo todo o resto intacto."""
    dados = json.loads(payload_bytes.decode("utf-8"))
    for query in dados.get("queries", []):
        commands = query.get("Query", {}).get("Commands", [])
        for cmd in commands:
            sqdsc = cmd.get("SemanticQueryDataShapeCommand")
            if not sqdsc:
                continue
            binding = sqdsc.get("Binding", {})
            reduction = binding.get("DataReduction", {})
            primary = reduction.get("Primary", {})
            if "Window" in primary:
                primary["Window"]["Count"] = novo_count
            elif "Top" in primary:
                primary["Top"]["Count"] = novo_count
    return json.dumps(dados).encode("utf-8")


def _extrair_nomes_colunas(payload_bytes: bytes) -> list[str]:
    """Lê o NativeReferenceName de cada Select do payload, na mesma ordem
    das projeções, para nomear as colunas do DataFrame de forma genérica."""
    dados = json.loads(payload_bytes.decode("utf-8"))
    query = dados["queries"][0]["Query"]["Commands"][0]["SemanticQueryDataShapeCommand"]["Query"]
    nomes = []
    for sel in query["Select"]:
        nome = sel.get("NativeReferenceName") or sel.get("Name") or "Coluna"
        nomes.append(nome)
    return nomes


def capturar_requisicao_e_resposta():
    """Abre o relatório num navegador headless, intercepta a requisição
    querydata correspondente ao CADIFA, aumenta a paginação e retorna
    (payload_da_requisicao, corpo_da_resposta)."""

    capturado = {"payload": None, "resposta": None}

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()

        def ao_interceptar_rota(route, request):
            if "querydata" not in request.url or capturado["resposta"] is not None:
                route.continue_()
                return

            post_data = request.post_data_buffer
            if not post_data or MARCADOR_ENTIDADE.encode() not in post_data:
                route.continue_()
                return

            # Guarda o payload original (para extrair nomes de coluna depois)
            # e envia a versão com Window.Count ajustado.
            capturado["payload"] = post_data
            novo_payload = _ajustar_window_count(post_data, PAGE_SIZE)
            route.continue_(post_data=novo_payload)

        def ao_receber_resposta(response):
            if "querydata" not in response.url or capturado["resposta"] is not None:
                return
            if capturado["payload"] is None:
                return
            try:
                capturado["resposta"] = response.body()
            except Exception:
                pass

        page.route("**/public/reports/querydata*", ao_interceptar_rota)
        page.on("response", ao_receber_resposta)

        print("Abrindo o relatório da Anvisa (headless)...")
        page.goto(REPORT_URL, timeout=TIMEOUT_MS)

        try:
            page.wait_for_function(
                "() => window.__cadifa_ok === true",
                timeout=1,
            )
        except Exception:
            pass

        # Aguarda até a resposta ser capturada (ou timeout)
        page.wait_for_timeout(1000)
        tentativas = TIMEOUT_MS // 1000
        while capturado["resposta"] is None and tentativas > 0:
            page.wait_for_timeout(1000)
            tentativas -= 1

        browser.close()

    if capturado["resposta"] is None or capturado["payload"] is None:
        print(
            "ERRO: não consegui capturar a requisição/resposta do painel.\n"
            "Possíveis causas:\n"
            "  - A Anvisa mudou a URL do endpoint (não é mais "
            "/public/reports/querydata).\n"
            "  - O relatório agora exige interação (clique/filtro) antes de "
            "carregar os dados.\n"
            "  - O tempo de carregamento foi maior que o timeout.\n"
            "Rode com headless=False (veja o comentário no código) para "
            "observar visualmente o que está acontecendo.",
            file=sys.stderr,
        )
        sys.exit(1)

    return capturado["payload"], capturado["resposta"]


def decodificar_dsr(resposta_bytes: bytes, nomes_colunas: list[str]) -> pd.DataFrame:
    """Decodifica o formato compactado DSR do Power BI em um DataFrame."""
    resposta = json.loads(resposta_bytes.decode("utf-8"))
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

    n_col_reais = len(linhas_decodificadas[0]) if linhas_decodificadas else len(nomes_colunas)
    if n_col_reais != len(nomes_colunas):
        # A estrutura do relatório mudou (coluna adicionada/removida).
        # Não travamos: geramos nomes genéricos para não perder dados.
        nomes_colunas = [f"Coluna_{i+1}" for i in range(n_col_reais)]
        print(
            f"AVISO: número de colunas retornadas ({n_col_reais}) difere do "
            f"esperado. Usando nomes genéricos - confira o Excel gerado."
        )

    df = pd.DataFrame(linhas_decodificadas, columns=nomes_colunas)

    # Tenta converter automaticamente qualquer coluna de data (epoch em ms)
    for col in df.columns:
        if "data" in col.lower() or "dt_" in col.lower():
            try:
                convertida = pd.to_datetime(df[col], unit="ms")
                df[col] = convertida.dt.strftime("%d/%m/%Y")
            except (ValueError, TypeError):
                pass

    return df


def main():
    payload, resposta = capturar_requisicao_e_resposta()
    nomes_colunas = _extrair_nomes_colunas(payload)

    df = decodificar_dsr(resposta, nomes_colunas)
    print(f"Total de linhas obtidas: {len(df)}")

    if len(df) == PAGE_SIZE:
        print("ATENÇÃO: o número de linhas retornado é igual ao solicitado (PAGE_SIZE).")
        print("Isso pode indicar que ainda existem mais linhas (limite atingido).")
        print("Aumente PAGE_SIZE no topo do script e rode novamente.")

    saida = "cadifa_completo.xlsx"
    df.to_excel(saida, index=False)
    print(f"Arquivo salvo: {saida}")


if __name__ == "__main__":
    main()
