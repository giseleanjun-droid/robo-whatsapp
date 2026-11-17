# -*- coding: utf-8 -*-
"""
ler_matriz.py - lê (SOMENTE LEITURA) a aba "📲 WHATSAPP" da planilha MATRIZ SLA + BACKLOG e os grupos
da planilha GRUPOS WHATSAPP, e grava mensagens.json para o enviar.mjs.

Para se a matriz não for de hoje (o robô das 04h não rodou): não manda mensagem velha.

Variáveis de ambiente:
  DRIVE_TOKEN   caminho do JSON da conta de serviço (segredo)
  PASTA_SLA     id da pasta "Bases Consolidadas" do SLA (segredo)
  SO_BASES      opcional: "GO-W-D001 GO-W-D009" -> só essas bases
  IGNORAR_DATA  "1" -> aceita matriz de outro dia (só para teste)
Saída: mensagens.json  [{base, uf, grupo, link, texto}]  e, se parar, motivo.txt
Linha DIRETORIA na GRUPOS WHATSAPP = também manda o resumo da diretoria (diretoria.py) nesse grupo.
"""
import io
import json
import os
import sys
import datetime as dt
import unicodedata
from pathlib import Path
from urllib.parse import quote
from zoneinfo import ZoneInfo

import pandas as pd
from google.oauth2 import service_account
from google.auth.transport.requests import AuthorizedSession

AQUI = Path(__file__).resolve().parent
ESCOPO = ["https://www.googleapis.com/auth/drive.readonly"]
PLANILHA = "application/vnd.google-apps.spreadsheet"
NOME_MATRIZ = os.environ.get("NOME_MATRIZ", "MATRIZ SLA + BACKLOG")
NOME_GRUPOS = os.environ.get("NOME_GRUPOS", "GRUPOS WHATSAPP")


def norm(c):
    c = unicodedata.normalize("NFD", str(c).lower())
    return "".join(ch for ch in c if ch.isascii() and ch.isalnum())


def parar(motivo):
    (AQUI / "motivo.txt").write_text(motivo, encoding="utf-8")
    print("PAROU:", motivo)
    sys.exit(3)


def achar(sessao, pasta, nome):
    q = f"'{pasta}' in parents and name = '{nome}' and trashed = false and mimeType = '{PLANILHA}'"
    r = sessao.get("https://www.googleapis.com/drive/v3/files",
                   params={"q": q, "fields": "files(id)", "supportsAllDrives": "true",
                           "includeItemsFromAllDrives": "true", "orderBy": "modifiedTime desc"}, timeout=60)
    r.raise_for_status()
    f = r.json().get("files", [])
    return f[0]["id"] if f else None


def aba(sessao, fid, nome, intervalo=None, cabecalho=1):
    url = f"https://docs.google.com/spreadsheets/d/{fid}/gviz/tq?tqx=out:csv&headers={cabecalho}&sheet={quote(nome)}"
    if intervalo:
        url += f"&range={intervalo}"
    r = sessao.get(url, timeout=300)
    r.raise_for_status()
    return pd.read_csv(io.BytesIO(r.content), dtype=str, keep_default_na=False, header=0 if cabecalho else None)


def main():
    cred = service_account.Credentials.from_service_account_file(os.environ["DRIVE_TOKEN"], scopes=ESCOPO)
    s = AuthorizedSession(cred)
    pasta = os.environ["PASTA_SLA"]
    fid = achar(s, pasta, NOME_MATRIZ)
    if not fid:
        parar(f"planilha '{NOME_MATRIZ}' não encontrada")

    # carimbo "SLA: 15/11 04:12 · Backlog: ..." na célula F1 da aba
    carimbo = " ".join(aba(s, fid, "📲 WHATSAPP", "A1:K1", cabecalho=0).astype(str).values.ravel())
    hoje = dt.datetime.now(ZoneInfo("America/Sao_Paulo")).strftime("%d/%m")
    if f"SLA: {hoje}" not in carimbo and os.environ.get("IGNORAR_DATA") != "1":
        parar(f"a matriz não é de hoje ({hoje}). Carimbo encontrado: {carimbo.strip()[:80] or 'vazio'}")

    msg = aba(s, fid, "📲 WHATSAPP", "A2:Z1000")
    msg.columns = [norm(c) for c in msg.columns]
    col_txt = next(c for c in msg.columns if c.startswith("mensagem"))

    grupos = {}
    gid = achar(s, pasta, NOME_GRUPOS)
    if gid:
        g = aba(s, gid, "GRUPOS")
        g.columns = [norm(c) for c in g.columns]
        cg = next((c for c in g.columns if "grupo" in c and "link" not in c), None)
        cl = next((c for c in g.columns if "link" in c), None)
        for r in g.to_dict("records"):
            b = str(r.get("base", "")).strip().upper()
            if b:
                grupos[b] = {"grupo": str(r.get(cg, "") or "").strip(), "link": str(r.get(cl, "") or "").strip()}

    so = set(os.environ.get("SO_BASES", "").upper().split())
    out = []
    for r in msg.to_dict("records"):
        base = str(r.get("base", "")).strip().upper()
        if not base or (so and base not in so):
            continue
        if grupos and base not in grupos:   # apagada da GRUPOS WHATSAPP = não envia (inativa, ou estado que não recebe)
            continue
        g = grupos.get(base, {})
        out.append({"base": base, "uf": base[:2], "grupo": g.get("grupo", ""), "link": g.get("link", ""),
                    "texto": str(r.get(col_txt, ""))})
    # resumo da DIRETORIA: linha "DIRETORIA" na planilha GRUPOS WHATSAPP (se falhar, as bases saem normalmente)
    g = grupos.get("DIRETORIA")
    if g and (g.get("grupo") or g.get("link")) and (not so or "DIRETORIA" in so):
        try:
            import diretoria
            out.insert(0, {"base": "DIRETORIA", "uf": "", "grupo": g["grupo"], "link": g["link"], "texto": diretoria.montar(s, pasta)})
            print("resumo da diretoria OK")
        except Exception as e:
            print("resumo da diretoria NÃO montado:", type(e).__name__)
    (AQUI / "mensagens.json").write_text(json.dumps(out, ensure_ascii=False), encoding="utf-8")
    com = sum(1 for x in out if x["grupo"] or x["link"])
    print(f"matriz de hoje OK: {len(out)} bases, {com} com grupo configurado")


if __name__ == "__main__":
    main()
