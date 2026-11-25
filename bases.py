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


def exportar_xlsx(svc, fid):
    """A planilha de cada base é um Google Sheets NATIVO (não um arquivo binário): get_media() não funciona nele
    (dá 403) - precisa exportar para xlsx."""
    buf = io.BytesIO()
    dl = MediaIoBaseDownload(buf, svc.files().export_media(fileId=fid, mimeType=XLSX))
    feito = False
    while not feito:
        _, feito = dl.next_chunk()
    return buf.getvalue()


def ler_anterior(bruto):
    """Lê o que a base já preencheu na planilha atual (antes de sobrescrever): telefone colado e resposta do
    cliente marcada, por Waybill. Sem isso, cada regravação automática apagaria o que a base digitou."""
    from openpyxl import load_workbook
    try:
        wb = load_workbook(io.BytesIO(bruto), read_only=True, data_only=True)
        if "🚨 BACKLOG PACOTES" not in wb.sheetnames:
            return {}
        ws = wb["🚨 BACKLOG PACOTES"]
        linhas = ws.iter_rows(values_only=True)
        next(linhas, None)   # linha 1: resumo
        cab = next(linhas, None)   # linha 2: cabeçalho
        if not cab:
            return {}
        idx = {nome: i for i, nome in enumerate(cab) if nome}
        j_wb, j_tel, j_resp = idx.get("Waybill"), idx.get("Telefone (cole aqui)"), idx.get("Resposta do cliente")
        if j_wb is None or (j_tel is None and j_resp is None):
            return {}
        out = {}
        for row in linhas:
            if j_wb >= len(row) or not row[j_wb]:
                continue
            tel = row[j_tel] if j_tel is not None and j_tel < len(row) else None
            resp = row[j_resp] if j_resp is not None and j_resp < len(row) else None
            if tel or resp:
                out[row[j_wb]] = {"telefone": str(tel).strip() if tel else None, "resposta": resp or None}
        return out
    except Exception:
        return {}


def aba_expedido(caminho, base, exped):
    """Acrescenta a aba "📦 EXPEDIDO GO-74" na planilha da base: o que saiu no último turno e o que saiu há +24h sem entrada."""
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill
    wb = load_workbook(caminho)
    ws = wb.create_sheet("📦 EXPEDIDO GO-74")
    b = (exped.get("bases") or {}).get(base) or {"n": 0, "lista": [], "pend": []}
    cinza, negrito = PatternFill("solid", fgColor="E5E7EB"), Font(bold=True)
    ws.append([f"🚚 SAÍRAM DO GO-74 PARA {base} · saídas {exped.get('quando', '')} · {b['n']} pacotes · atualizado {exped.get('gerado', '')}"])
    ws.append(["Pacote", "Cidade", "Bairro", "Veículo", "Lacre"])
    for c in ws[2]:
        c.font, c.fill = negrito, cinza
    for l in sorted(b["lista"], key=lambda x: (str(x[1]), str(x[2]))):
        ws.append(l)
    ws.append([])
    ws.append([f"⏳ SAÍRAM DO GO-74 HÁ MAIS DE 24H E NÃO DERAM ENTRADA · {len(b['pend'])} pacotes · se já está na base, dê entrada; se não chegou, registre na Exceção de Transferência"])
    t2 = ws.max_row
    ws.append(["Pacote", "Cidade", "Bairro", "Saída do GO-74", "Dias"])
    for c in ws[t2 + 1]:
        c.font, c.fill = negrito, cinza
    for l in sorted(b["pend"], key=lambda x: -(x[4] or 0)):
        ws.append(l)
    for r in (1, t2):
        ws.cell(r, 1).font = Font(bold=True, size=12)
    for col, w in zip("ABCDE", (22, 24, 28, 18, 18)):
        ws.column_dimensions[col].width = w
    wb.save(caminho)


def main():
    import datetime as dt, time
    from zoneinfo import ZoneInfo
    svc = servico()
    # De manhã, espera o robô do SRM terminar a rodada das 04h (24/11 ele terminou 05h32 e as planilhas
    # foram gravadas às 05h17 com os dados de ontem). Até 60 min, conferindo a cada 3 min.
    fim = time.time() + 60 * 60
    while True:
        arquivos = listar(svc, os.environ["PASTA_SLA"])
        dados = arquivos.get("_DADOS_BASES.bin")
        if not dados:
            sys.exit("_DADOS_BASES.bin não existe na pasta (rode o CriarPlanilhasBases.gs)")
        bruto = baixar(svc, dados["id"])
        if len(bruto) < 100:
            sys.exit("_DADOS_BASES.bin ainda vazio: espere a próxima rodada do robô (04h, 14h ou 19h)")
        d = pickle.loads(gzip.decompress(bruto))
        S, B = d["S"], d["B"]
        agora = dt.datetime.now(ZoneInfo("America/Sao_Paulo"))
        if agora.hour >= 10 or S["hoje"].date() >= agora.date() or time.time() > fim:
            break
        print(f"dados ainda de {S['hoje']:%d/%m %H:%M}: esperando o robô das 04h terminar…", flush=True)
        time.sleep(180)
    print(f"dados do robô: SLA {S['hoje']:%d/%m %H:%M} · Backlog {B['hoje']:%d/%m %H:%M} (arquivo de {dados['modifiedTime'][:16]} UTC)")

    # listas da expedição do GO-74 (Painel HUB) para a aba "📦 EXPEDIDO GO-74"
    exped = None
    if arquivos.get("_EXPEDICAO_LISTAS.json"):
        try:
            import json
            exped = json.loads(baixar(svc, arquivos["_EXPEDICAO_LISTAS.json"]["id"]).decode("utf-8"))
            print(f"expedição GO-74: listas de {exped.get('gerado')} ({len(exped.get('bases', {}))} bases)")
        except Exception as e:
            print(f"expedição GO-74: listas não lidas ({e})")

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
                s = servico()   # um cliente por thread (o httplib2 não é thread-safe)
                anterior = ler_anterior(exportar_xlsx(s, fid))
                p = planilha_base.gerar(base, S, B, tmp / f"{base}.xlsx", anterior)
                if exped is not None:
                    aba_expedido(p, base, exped)
                s.files().update(fileId=fid, media_body=MediaFileUpload(str(p), mimetype=XLSX, resumable=False),
                                 supportsAllDrives=True).execute()
                p.unlink(missing_ok=True)
                return base, "ok"
            except Exception as e:
                erro = str(e)[:300]
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
