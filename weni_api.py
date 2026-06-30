#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Integracao com a API da Weni (plataforma RapidPro/Temba)
========================================================
Baixa conversas (messages) e contatos (contacts) direto da API da Weni e salva
na pasta de entradas, no MESMO formato dos arquivos exportados manualmente, para
o funil_weni.py processar em seguida sem nenhuma alteracao.

  https://flows.weni.ai/api/v2/messages.json
  https://flows.weni.ai/api/v2/contacts.json

----------------------------------------------------------------------------
COMO CONFIGURAR (voce so precisa mexer aqui em baixo, uma vez):

1. TOKEN: pegue o seu token no API Explorer da Weni
   (https://dash.weni.ai/api/flows/api/v2/explorer) — fica escrito
   "Authorization: Token XXXXXXXX" no topo da pagina. Copie a parte XXXXXXXX.

   NAO escreva o token aqui no codigo. Em vez disso, crie uma variavel de
   ambiente chamada WENI_TOKEN (o passo a passo esta no README, secao API).
   O script le ela sozinho. Se preferir o jeito rapido (menos seguro), da para
   colar o token na linha TOKEN_FIXO abaixo.

2. Pronto. Rode:  python weni_api.py --output entradas
----------------------------------------------------------------------------
"""

import argparse
import os
import sys
import time
from datetime import datetime, timedelta

try:
    import requests
    import pandas as pd
except ImportError as e:
    sys.exit("Dependencia faltando: %s\nRode:  pip install -r requirements.txt" % e)

# ============================================================================
# CONFIGURACAO
# ============================================================================
BASE_URL = "https://flows.weni.ai/api/v2"
MESSAGES_URL = BASE_URL + "/messages.json"
CONTACTS_URL = BASE_URL + "/contacts.json"

# A Weni pode dar UM token unico (serve para os dois) ou DOIS tokens separados
# (um para conversas/messages e outro para contatos/contacts). Suporta os dois casos:
#
#  - Token unico:   defina a variavel de ambiente WENI_TOKEN
#  - Dois tokens:   defina WENI_TOKEN_MESSAGES e WENI_TOKEN_CONTACTS
#
# Para teste rapido, da para colar direto aqui (cuidado: fica exposto no arquivo).
TOKEN_FIXO = ""                 # token unico (se a Weni te deu so um)
TOKEN_FIXO_MESSAGES = ""        # token das CONVERSAS (se forem dois)
TOKEN_FIXO_CONTACTS = ""        # token dos CONTATOS (se forem dois)

# Quantos dias de conversas baixar para tras (usado apenas se voce NAO fixar
# uma data manualmente la embaixo, na funcao main). A base de contatos vem
# sempre completa, sem filtro de data.
DIAS_DE_HISTORICO = 45

# Pausa entre paginas (segundos). O limite da Weni e 2500 chamadas/hora.
PAUSA_ENTRE_PAGINAS = 0.3


# ============================================================================
# FUNCOES
# ============================================================================
def _resolve(especifico_fixo, especifico_env):
    """Resolve um token: usa o especifico se existir, senao cai no token unico."""
    return (
        especifico_fixo.strip()
        or os.environ.get(especifico_env, "").strip()
        or TOKEN_FIXO.strip()
        or os.environ.get("WENI_TOKEN", "").strip()
    )


def get_tokens():
    tk_msg = _resolve(TOKEN_FIXO_MESSAGES, "WENI_TOKEN_MESSAGES")
    tk_con = _resolve(TOKEN_FIXO_CONTACTS, "WENI_TOKEN_CONTACTS")
    if not tk_msg or not tk_con:
        sys.exit(
            "Token(s) nao encontrado(s).\n"
            "Se a Weni te deu DOIS tokens, defina as duas variaveis:\n"
            '   setx WENI_TOKEN_MESSAGES "token_das_conversas"\n'
            '   setx WENI_TOKEN_CONTACTS "token_dos_contatos"\n'
            "Se for UM token unico, defina:\n"
            '   setx WENI_TOKEN "seu_token"\n'
            "Depois feche e abra o PowerShell e rode de novo."
        )
    return tk_msg, tk_con


def fetch_all(url, token, params=None, rotulo=""):
    """Percorre todas as paginas de um endpoint RapidPro usando o cursor 'next'."""
    headers = {"Authorization": "Token %s" % token, "Accept": "application/json"}
    resultados = []
    pagina = 0
    proxima = url
    primeira_params = dict(params or {})
    tentativas_erro = 0
    MAX_TENTATIVAS = 8

    while proxima:
        pagina += 1
        try:
            if pagina == 1:
                resp = requests.get(proxima, headers=headers, params=primeira_params, timeout=60)
            else:
                # nas proximas paginas a URL ja vem completa com o cursor
                resp = requests.get(proxima, headers=headers, timeout=60)
        except requests.RequestException as e:
            tentativas_erro += 1
            if tentativas_erro > MAX_TENTATIVAS:
                sys.exit("Erro de conexao ao buscar %s apos %d tentativas: %s" % (rotulo, MAX_TENTATIVAS, e))
            espera = min(10 * tentativas_erro, 60)
            print("   erro de conexao em %s, tentativa %d/%d, aguardando %ds..." % (rotulo, tentativas_erro, MAX_TENTATIVAS, espera))
            time.sleep(espera)
            pagina -= 1
            continue

        if resp.status_code == 429:  # rate limit
            espera = int(resp.headers.get("Retry-After", 60))
            print("   limite de chamadas atingido, aguardando %ds..." % espera)
            time.sleep(espera)
            pagina -= 1
            continue
        if resp.status_code in (500, 502, 503, 504):  # erro temporario do servidor da Weni
            tentativas_erro += 1
            if tentativas_erro > MAX_TENTATIVAS:
                sys.exit("Erro %d persistente ao buscar %s apos %d tentativas." % (resp.status_code, rotulo, MAX_TENTATIVAS))
            espera = min(15 * tentativas_erro, 120)
            print("   erro %d (temporario) em %s, tentativa %d/%d, aguardando %ds..." % (resp.status_code, rotulo, tentativas_erro, MAX_TENTATIVAS, espera))
            time.sleep(espera)
            pagina -= 1
            continue
        if resp.status_code == 403:
            sys.exit("Acesso negado (403) em %s. O token tem permissao de leitura desse recurso?" % rotulo)
        if resp.status_code != 200:
            sys.exit("Erro %d ao buscar %s: %s" % (resp.status_code, rotulo, resp.text[:300]))

        tentativas_erro = 0  # reset apos sucesso
        data = resp.json()
        lote = data.get("results", [])
        resultados.extend(lote)
        print("   %s: pagina %d (+%d, total %d)" % (rotulo, pagina, len(lote), len(resultados)))
        proxima = data.get("next")
        if proxima:
            time.sleep(PAUSA_ENTRE_PAGINAS)

    return resultados


def normaliza_mensagens(msgs):
    """Converte o JSON de messages no mesmo formato do export manual (colunas)."""
    linhas = []
    for m in msgs:
        contact = m.get("contact") or {}
        channel = m.get("channel") or {}
        urn = m.get("urn") or ""
        if urn.startswith("tel:"):
            urn = urn[4:].lstrip("+")
        linhas.append({
            "Date": m.get("created_on"),
            "Direction": "IN" if m.get("direction") == "in" else "OUT",
            "Text": m.get("text") or "",
            "Status": m.get("status") or "",
            "Channel": channel.get("name") or "",
            "Contact UUID": contact.get("uuid") or "",
            "Name": contact.get("name") or "",
            "URN": urn,
        })
    df = pd.DataFrame(linhas)
    if not df.empty:
        df["Date"] = pd.to_datetime(df["Date"]).dt.tz_localize(None)
    return df


# Mapeamento de chaves da API (minusculas) para o formato do export manual
FIELD_MAP = {
    "analista":    "Field:Analista",
    "cnpj":        "Field:CNPJ",
    "uf":          "Field:UF",
    "estado":      "Field:Estado",
    "cidade":      "Field:Cidade",
    "regional":    "Field:Regional",
    "razao social":"Field:Razao Social",
    "razaosocial": "Field:Razao Social",
    "nome fantasia":"Field:Nome Fantasia",
    "nomefantasia": "Field:Nome Fantasia",
    "hunter":      "Field:Hunter",
    "sellers":     "Field:Sellers",
    "prioridade de intencao de compra": "Field:Prioridade de intencao de compra",
    "score de intencao de compra":      "Field:Score de intencao de compra",
    "purchase_intent_priority":         "Field:Prioridade de intencao de compra",
    "purchase_intent_score":            "Field:Score de intencao de compra",
}


def _map_field(chave):
    """Normaliza chave da API para o formato Field:NomeOriginal."""
    lower = chave.lower().strip()
    if lower in FIELD_MAP:
        return FIELD_MAP[lower]
    # para campos nao mapeados, preserva com prefixo Field: e capitaliza
    return "Field:%s" % chave


def normaliza_contatos(cons):
    """Converte o JSON de contacts no mesmo formato do export manual (Field:...)."""
    linhas = []
    for c in cons:
        fields = c.get("fields") or {}
        urns = c.get("urns") or []
        tel = ""
        for u in urns:
            if u.startswith("tel:"):
                tel = u[4:].lstrip("+"); break
        linha = {
            "Contact UUID": c.get("uuid") or "",
            "Name": c.get("name") or "",
            "URN": tel,
        }
        # cada campo personalizado da Weni vira uma coluna "Field:<nome>"
        for chave, valor in fields.items():
            linha[_map_field(chave)] = valor
        linhas.append(linha)
    return pd.DataFrame(linhas)


def main():
    ap = argparse.ArgumentParser(description="Baixa conversas e contatos da API da Weni.")
    ap.add_argument("--output", default="./entradas", help="pasta onde salvar os .xlsx")
    ap.add_argument("--dias", type=int, default=DIAS_DE_HISTORICO,
                    help="quantos dias de conversas baixar (padrao %d). Ignorado, pois ha uma data fixa definida no codigo." % DIAS_DE_HISTORICO)
    args = ap.parse_args()

    tk_msg, tk_con = get_tokens()
    os.makedirs(args.output, exist_ok=True)

    # Data fixa de inicio: busca tudo desde 05/05/2026 ate o momento atual.
    desde = "2026-05-05T00:00:00.000000Z"

    print("Baixando CONTATOS da Weni...")
    contatos = fetch_all(CONTACTS_URL, tk_con, rotulo="contatos")
    df_con = normaliza_contatos(contatos)

    print("Baixando CONVERSAS da Weni (desde 05/05/2026)...")
    mensagens = fetch_all(MESSAGES_URL, tk_msg, params={"after": desde}, rotulo="conversas")
    df_msg = normaliza_mensagens(mensagens)

    # nomes iguais aos exports manuais, para o funil_weni achar por prefixo
    msg_path = os.path.join(args.output, "message_export_api.xlsx")
    con_path = os.path.join(args.output, "contact_export_api.xlsx")
    df_msg.to_excel(msg_path, index=False)
    df_con.to_excel(con_path, index=False)

    print("\nOK")
    print("   Conversas: %d mensagens -> %s" % (len(df_msg), os.path.basename(msg_path)))
    print("   Contatos:  %d contatos  -> %s" % (len(df_con), os.path.basename(con_path)))
    print("\nAgora rode:  python funil_weni.py --input %s --output saidas" % args.output)


if __name__ == "__main__":
    main()
