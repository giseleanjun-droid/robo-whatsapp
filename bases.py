# -*- coding: utf-8 -*-
"""
bases.py - grava a planilha de CADA base ("<BASE> - SLA + BACKLOG") com os dados que o robô das 04h/14h/19h
deixou no arquivo _DADOS_BASES.bin (pasta "PLANILHAS DAS BASES (SLA + BACKLOG)" dentro da pasta do SLA).

Roda no repositório PÚBLICO (minutos ilimitados). O log só mostra contagens.
Variáveis: DRIVE_TOKEN (JSON da conta de serviço), PASTA_SLA, SO_BASES (opcional, para teste).
"""
import gzip
import io
import os
import pickle
import sys
import tempfile
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload

AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI / "MATRIZ"))
import planilha_base   # noqa: E402

NOME_PASTA = "PLANILHAS DAS BASES (SLA + BACKLOG)"
SUFIXO = " - SLA + BACKLOG"
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def servico():
    cred = service_account.Credentials.from_service_account_file(
        os.environ["DRIVE_TOKEN"], scopes=["https://www.googleapis.com/auth/drive"])
    return build("drive", "v3", credentials=cred, cache_discovery=False)


def listar(svc, pasta):
    q = f"'{pasta}' in parents and name = '{NOME_PASTA}' and trashed = false and mimeType = 'application/vnd.google-apps.folder'"
    p = svc.files().list(q=q, fields="files(id)", supportsAllDrives=True, includeItemsFromAllDrives=True).execute()["files"]
    if not p:
        sys.exit(f"pasta '{NOME_PASTA}' não existe (rode o CriarPlanilhasBases.gs)")
    arquivos, token = {}, None
    while True:
        r = svc.files().list(q=f"'{p[0]['id']}' in parents and trashed = false", fields="nextPageToken,files(id,name,modifiedTime)",
                             pageSize=1000, pageToken=token, supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
        for f in r["files"]:
            arquivos[f["name"]] = f
        token = r.get("nextPageToken")
        if not token:
            return arquivos


def baixar(svc, fid):
    buf = io.BytesIO()
    dl = MediaIoBaseDownload(buf, svc.files().get_media(fileId=fid, supportsAllDrives=True))
    feito = False
    while not feito:
        _, feito = dl.next_chunk()
    return buf.getvalue()


def main():
    svc = servico()
    arquivos = listar(svc, os.environ["PASTA_SLA"])
    dados = arquivos.get("_DADOS_BASES.bin")
    if not dados:
        sys.exit("_DADOS_BASES.bin não existe na pasta (rode o CriarPlanilhasBases.gs)")
    bruto = baixar(svc, dados["id"])
    if len(bruto) < 100:
        sys.exit("_DADOS_BASES.bin ainda vazio: espere a próxima rodada do robô (04h, 14h ou 19h)")
    d = pickle.loads(gzip.decompress(bruto))
    S, B = d["S"], d["B"]
    print(f"dados do robô: SLA {S['hoje']:%d/%m %H:%M} · Backlog {B['hoje']:%d/%m %H:%M} (arquivo de {dados['modifiedTime'][:16]} UTC)")

    so = set(os.environ.get("SO_BASES", "").upper().split())
    alvo = {n[:-len(SUFIXO)].strip().upper(): f["id"] for n, f in arquivos.items() if n.endswith(SUFIXO)}
    if so:
        alvo = {b: i for b, i in alvo.items() if b in so}
    todas = set(S["bases"]["base"]) | set(B["bases"]["base"])
    sem_planilha = sorted(todas - set(alvo)) if not so else []
    tmp = Path(tempfile.mkdtemp(prefix="bases_"))

    def uma(item):
        base, fid = item
        for n in range(1, 4):
            try:
                p = planilha_base.gerar(base, S, B, tmp / f"{base}.xlsx")
                s = servico()   # um cliente por thread (o httplib2 não é thread-safe)
                s.files().update(fileId=fid, media_body=MediaFileUpload(str(p), mimetype=XLSX, resumable=False),
                                 supportsAllDrives=True).execute()
                p.unlink(missing_ok=True)
                return base, "ok"
            except Exception as e:
                erro = str(e)[:120]
                time.sleep(5 * n)
        return base, "falhou: " + erro

    t0 = time.time()
    with ThreadPoolExecutor(max_workers=6) as ex:
        res = list(ex.map(uma, sorted(alvo.items())))
    ok = [b for b, s in res if s == "ok"]
    falhas = [(b, s) for b, s in res if s != "ok"]
    print(f"planilhas atualizadas: {len(ok)} de {len(alvo)} em {time.time() - t0:.0f}s")
    for b, s in falhas:
        print(f"  {b}: {s}")
    if sem_planilha:
        print(f"bases SEM planilha criada ({len(sem_planilha)}): {' '.join(sem_planilha)} -> rode o CriarPlanilhasBases.gs de novo")
    sys.exit(1 if falhas else 0)


if __name__ == "__main__":
    main()
