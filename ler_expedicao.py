# -*- coding: utf-8 -*-
"""
ler_expedicao.py - lê (SOMENTE LEITURA) o _EXPEDICAO_MSG.json que o projeto "Painel HUB" do Apps Script grava
na pasta do SLA (o que o GO-74 expediu para cada base, por bairro) e os grupos da planilha GRUPOS WHATSAPP.
Grava mensagens.json para o enviar.mjs.

Para (motivo.txt) se o arquivo não for do turno de agora: não manda mensagem velha.
  TURNO   manha | tarde   (padrão: manhã antes do meio-dia)
  SO_BASES opcional: só essas bases
"""
import datetime as dt
import json
import os
import sys
from pathlib import Path
from zoneinfo import ZoneInfo

from google.oauth2 import service_account
from google.auth.transport.requests import AuthorizedSession

from ler_matriz import achar, aba, norm, parar, ESCOPO, NOME_GRUPOS

AQUI = Path(__file__).resolve().parent
NOME_MSG = "_EXPEDICAO_MSG.json"


def main():
    cred = service_account.Credentials.from_service_account_file(os.environ["DRIVE_TOKEN"], scopes=ESCOPO)
    s = AuthorizedSession(cred)
    pasta = os.environ["PASTA_SLA"]
    agora = dt.datetime.now(ZoneInfo("America/Sao_Paulo"))
    turno = (os.environ.get("TURNO") or ("manha" if agora.hour < 12 else "tarde")).lower()

    q = f"'{pasta}' in parents and name = '{NOME_MSG}' and trashed = false"
    r = s.get("https://www.googleapis.com/drive/v3/files", params={"q": q, "fields": "files(id)", "supportsAllDrives": "true",
                                                                   "includeItemsFromAllDrives": "true"}, timeout=60)
    r.raise_for_status()
    f = r.json().get("files", [])
    if not f:
        parar(f"{NOME_MSG} não existe na pasta do SLA (rode hubInstalar no projeto Painel HUB)")
    m = s.get(f"https://www.googleapis.com/drive/v3/files/{f[0]['id']}", params={"alt": "media", "supportsAllDrives": "true"}, timeout=120).json()
    hoje = agora.strftime("%Y-%m-%d")
    if m.get("dia") != hoje or m.get("turno") != turno:
        parar(f"a expedição do GO-74 não é deste turno ({hoje} {turno}). Arquivo: {m.get('dia')} {m.get('turno')} gerado {m.get('gerado')}")

    grupos = {}
    gid = achar(s, pasta, NOME_GRUPOS)
    if gid:
        g = aba(s, gid, "GRUPOS")
        g.columns = [norm(c) for c in g.columns]
        cg = next((c for c in g.columns if "grupo" in c and "link" not in c), None)
        cl = next((c for c in g.columns if "link" in c), None)
        for row in g.to_dict("records"):
            b = str(row.get("base", "")).strip().upper()
            if b:
                grupos[b] = {"grupo": str(row.get(cg, "") or "").strip(), "link": str(row.get(cl, "") or "").strip()}

    so = set(os.environ.get("SO_BASES", "").upper().split())
    out = []
    for it in m.get("itens", []):
        base = str(it.get("base", "")).strip().upper()
        if not base or (so and base not in so) or not it.get("pacotes"):
            continue
        if grupos and base not in grupos:
            continue
        g = grupos.get(base, {})
        out.append({"base": base, "uf": base[:2], "grupo": g.get("grupo", ""), "link": g.get("link", ""), "texto": it["texto"]})
    (AQUI / "mensagens.json").write_text(json.dumps(out, ensure_ascii=False), encoding="utf-8")
    print(f"expedição GO-74 ({m.get('turno')} {m.get('dia')}): {len(out)} bases com grupo · {m.get('pacotes')} pacotes no total")


if __name__ == "__main__":
    main()
