# -*- coding: utf-8 -*-
"""
ler_expedicao.py - lê (SOMENTE LEITURA) o _EXPEDICAO_MSG.json que o projeto "Painel HUB" do Apps Script grava
na pasta do SLA (o que o GO-74 expediu para cada base, por bairro) e os grupos da planilha GRUPOS WHATSAPP.
Grava mensagens.json para o enviar.mjs.

Para (motivo.txt) se o arquivo não for do turno de agora: não manda mensagem velha.
  TURNO   manha | tarde | agora   (padrão: manhã antes do meio-dia)
            agora = pré-alerta a cada 30 min (_EXPEDICAO_AGORA.json): sem trava de dia/turno, só confere
            se o arquivo é recente; se não tem base com novidade, não é erro (grava sem_novidade.txt e sai)
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
NOME_AGORA = "_EXPEDICAO_AGORA.json"


def sem_novidade(motivo):
    (AQUI / "sem_novidade.txt").write_text(motivo, encoding="utf-8")
    print("PRÉ-ALERTA:", motivo)


def main():
    cred = service_account.Credentials.from_service_account_file(os.environ["DRIVE_TOKEN"], scopes=ESCOPO)
    s = AuthorizedSession(cred)
    pasta = os.environ["PASTA_SLA"]
    agora = dt.datetime.now(ZoneInfo("America/Sao_Paulo"))
    turno = (os.environ.get("TURNO") or ("manha" if agora.hour < 12 else "tarde")).lower()

    nome_arquivo = NOME_AGORA if turno == "agora" else NOME_MSG
    q = f"'{pasta}' in parents and name = '{nome_arquivo}' and trashed = false"
    r = s.get("https://www.googleapis.com/drive/v3/files", params={"q": q, "fields": "files(id)", "supportsAllDrives": "true",
                                                                   "includeItemsFromAllDrives": "true"}, timeout=60)
    r.raise_for_status()
    f = r.json().get("files", [])
    if not f:
        if turno == "agora":
            return sem_novidade(f"{NOME_AGORA} ainda não existe (rode hubInstalar no projeto Painel HUB)")
        parar(f"{NOME_MSG} não existe na pasta do SLA (rode hubInstalar no projeto Painel HUB)")
    m = s.get(f"https://www.googleapis.com/drive/v3/files/{f[0]['id']}", params={"alt": "media", "supportsAllDrives": "true"}, timeout=120).json()
    if turno == "agora":
        try:
            gerado = dt.datetime.strptime(m.get("gerado", ""), "%Y-%m-%d %H:%M").replace(tzinfo=ZoneInfo("America/Sao_Paulo"))
            idade_min = (agora - gerado).total_seconds() / 60
        except ValueError:
            idade_min = 9999
        if idade_min > 90:
            return sem_novidade(f"pré-alerta desatualizado (gerado {m.get('gerado')}, {idade_min:.0f} min atrás)")
        # dedup: se o Painel HUB ainda não rodou de novo desde o último disparo (ex.: alguém testou na mão
        # entre dois ciclos automáticos do Despertador), o arquivo é o MESMO - sem isso manda a msg repetida.
        marca = AQUI / "agora_anterior.txt"
        chave = f"{m.get('ini')}|{m.get('fim')}"
        if marca.exists() and marca.read_text(encoding="utf-8").strip() == chave:
            return sem_novidade(f"já enviado este lote (ini {m.get('ini')} · fim {m.get('fim')})")
        (AQUI / "agora_marca.txt").write_text(chave, encoding="utf-8")
    else:
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
        if not base or (so and base not in so):   # o Painel HUB já só manda base com expedição ou pendente de 3+ dias
            continue
        if grupos and base not in grupos:
            continue
        g = grupos.get(base, {})
        out.append({"base": base, "uf": base[:2], "grupo": g.get("grupo", ""), "link": g.get("link", ""), "texto": it["texto"]})
    if turno == "agora" and not out:
        return sem_novidade("nenhuma base com caminhão novo neste ciclo")
    (AQUI / "mensagens.json").write_text(json.dumps(out, ensure_ascii=False), encoding="utf-8")
    if turno == "agora":
        print(f"pré-alerta de carga: {len(out)} base(s) com novidade (gerado {m.get('gerado')})")
    else:
        print(f"expedição GO-74 ({m.get('turno')} {m.get('dia')}): {len(out)} bases com grupo · {m.get('pacotes')} pacotes no total")


if __name__ == "__main__":
    main()
