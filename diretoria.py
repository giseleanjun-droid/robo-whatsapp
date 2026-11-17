# -*- coding: utf-8 -*-
"""
diretoria.py - resumo diário para a DIRETORIA (as mesmas contas do Painel da Diretoria da TV).

Fontes (SOMENTE LEITURA, na pasta do SLA):
  _DASH_EXTRA.json ........... robô: SLA de ontem/mês por base, backlog
  _EVOLUCAO_MENSAL_DADOS ..... robô: histórico diário/mensal, causas e perdas oficiais

Envio: na planilha GRUPOS WHATSAPP, crie uma linha com BASE = DIRETORIA e o grupo (nome ou link)
ou o TELEFONE da pessoa com DDD (mensagem direta).
O link do painel vai no topo: o dirTestar do Apps Script grava em _PAINEL_DIRETORIA.txt (pasta do SLA).
Nada de dados aqui no código: tudo é lido do Drive na hora (repositório público).
"""
import io
import datetime as dt
from urllib.parse import quote
from zoneinfo import ZoneInfo

import pandas as pd

UFS = ["GO", "DF", "TO", "MT", "MS", "RO", "AC"]
META, PERIGO = 98.0, 95.0
MES = ["", "jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]
SEM = ["seg", "ter", "qua", "qui", "sex", "sáb", "dom"]
CAUSA_GRUPO = {1: "Transporte", 2: "HUB", 3: "Transporte", 4: "HUB", 5: "Perda", 6: "Transporte", 7: "Cliente",
               8: "Base", 9: "Base", 10: "Motorista", 11: "Base", 0: "Outros"}
CAUSA_NOME = {1: "Chegou vencido na base", 2: "Não saiu do HUB", 3: "Expedido e não chegou na base", 4: "Erro de triagem",
              5: "Perda / extravio", 6: "Chegou em cima do prazo", 7: "Insucesso com ocorrência", 8: "Parado na base",
              9: "Saiu tarde para rota", 10: "Em rota e não entregue a tempo", 11: "Entregue com atraso sem saída", 0: "Sem classificação"}


def _uf(base, uf=""):
    b = str(base or "").strip().upper()
    if b == "PA-A-PVH01":
        return "RO"
    if b == "PA-A-RBR01":
        return "AC"
    return b[:2] if len(b) > 2 and b[2] in "- " else str(uf or "").strip().upper()


def _pct(ok, tot):
    return 100.0 * ok / tot if tot else None


def _p(v, c=1):
    return "—" if v is None else f"{v:.{c}f}%".replace(".", ",")


def _pp(v):
    return "—" if v is None else f"{abs(v):.1f} pp".replace(".", ",")


def _n(v):
    return f"{round(v or 0):,}".replace(",", ".")


def _r(v):
    v = v or 0
    if abs(v) >= 1e6:
        return f"R$ {v / 1e6:.2f} mi".replace(".", ",")
    if abs(v) >= 1e3:
        return f"R$ {v / 1e3:.1f} mil".replace(".", ",")
    return f"R$ {_n(v)}"


def _bola(v):
    return "⚪" if v is None else ("🟢" if v >= META else ("🟡" if v >= PERIGO else "🔴"))


def _seta(d):
    return "" if d is None else (f"▲ {_pp(d)}" if d > 0.05 else (f"▼ {_pp(d)}" if d < -0.05 else "= 0 pp"))


def _arquivo(sessao, pasta, nome, planilha):
    mime = "and mimeType = 'application/vnd.google-apps.spreadsheet'" if planilha else ""
    q = f"'{pasta}' in parents and name = '{nome}' and trashed = false {mime}"
    r = sessao.get("https://www.googleapis.com/drive/v3/files", params={"q": q, "fields": "files(id)", "supportsAllDrives": "true",
                   "includeItemsFromAllDrives": "true", "orderBy": "modifiedTime desc"}, timeout=60)
    r.raise_for_status()
    f = r.json().get("files", [])
    return f[0]["id"] if f else None


def _aba(sessao, fid, nome):
    url = f"https://docs.google.com/spreadsheets/d/{fid}/gviz/tq?tqx=out:csv&headers=1&sheet={quote(nome)}"
    r = sessao.get(url, timeout=300)
    r.raise_for_status()
    return pd.read_csv(io.BytesIO(r.content))


def _tab(t):
    return pd.DataFrame(t["d"], columns=t["c"]) if t and t.get("c") else pd.DataFrame()


def link_painel(sessao, pasta):
    """Link do Painel da Diretoria, gravado pelo dirTestar do Apps Script em _PAINEL_DIRETORIA.txt."""
    fid = _arquivo(sessao, pasta, "_PAINEL_DIRETORIA.txt", planilha=False)
    if not fid:
        return ""
    r = sessao.get(f"https://www.googleapis.com/drive/v3/files/{fid}", params={"alt": "media", "supportsAllDrives": "true"}, timeout=60)
    return r.text.strip() if r.ok and r.text.strip().startswith("https://") else ""


def montar(sessao, pasta, painel=None):
    """Devolve o texto do resumo (WhatsApp) ou levanta erro se faltar fonte."""
    if painel is None:
        painel = link_painel(sessao, pasta)
    fid = _arquivo(sessao, pasta, "_DASH_EXTRA.json", planilha=False)
    if not fid:
        raise RuntimeError("_DASH_EXTRA.json não encontrado")
    r = sessao.get(f"https://www.googleapis.com/drive/v3/files/{fid}", params={"alt": "media", "supportsAllDrives": "true"}, timeout=120)
    r.raise_for_status()
    X = r.json()
    eid = _arquivo(sessao, pasta, "_EVOLUCAO_MENSAL_DADOS", planilha=True)
    if not eid:
        raise RuntimeError("_EVOLUCAO_MENSAL_DADOS não encontrado")
    diario, mensal, causas, perdas = (_aba(sessao, eid, n) for n in ("DIARIO", "MENSAL", "CAUSAS", "PERDAS"))
    plano = None
    pid = _arquivo(sessao, pasta, "_PLANO_ACAO.json", planilha=False)
    if pid:
        rp = sessao.get(f"https://www.googleapis.com/drive/v3/files/{pid}", params={"alt": "media", "supportsAllDrives": "true"}, timeout=60)
        if rp.ok:
            try:
                plano = rp.json().get("acoes", [])
            except ValueError:
                plano = None
    return texto(X, diario, mensal, causas, perdas, painel, plano=plano)


def texto(X, diario, mensal, causas, perdas, painel="", agora=None, plano=None):
    agora = agora or dt.datetime.now(ZoneInfo("America/Sao_Paulo"))
    hoje = agora.date()
    # "hoje" = dia da última rodada do robô (sla_ref "17/11 04:46"), se for dos últimos 3 dias
    try:
        dd, mm = str(X.get("sla_ref", ""))[:5].split("/")
        ref = dt.date(hoje.year - (1 if int(mm) > hoje.month + 1 else 0), int(mm), int(dd))
        if hoje - dt.timedelta(days=3) <= ref <= hoje:
            hoje = ref
    except ValueError:
        pass
    ontem, sp = hoje - dt.timedelta(days=1), hoje - dt.timedelta(days=8)
    ano, mes = hoje.year, hoje.month
    mes_ant, ano_ant = (12, ano - 1) if mes == 1 else (mes - 1, ano)

    b = _tab(X.get("bases"))
    bl = _tab(X.get("bl"))
    if b.empty:
        raise RuntimeError("sem bases no _DASH_EXTRA.json")
    b["UF"] = [_uf(x, u) for x, u in zip(b["base"], b["uf"])]
    b = b[b["UF"].isin(UFS)]
    num = ["dem_ontem", "suc_ontem", "dem_mes", "suc_mes", "demb_mes", "sucb_mes"]
    b[num] = b[num].apply(pd.to_numeric, errors="coerce").fillna(0)
    if not bl.empty:
        bl["UF"] = [_uf(x, u) for x, u in zip(bl["base"], bl["uf"])]
        bl = bl[bl["UF"].isin(UFS)]
        bl[["total", "valor", "vencidos", "extravio"]] = bl[["total", "valor", "vencidos", "extravio"]].apply(pd.to_numeric, errors="coerce").fillna(0)
    R = b[num].sum()
    sla_ont, sla_mes, sla_bm = _pct(R.suc_ontem, R.dem_ontem), _pct(R.suc_mes, R.dem_mes), _pct(R.sucb_mes, R.demb_mes)

    # histórico diário (data de prazo)
    d = diario.copy()
    d["UF"] = [_uf(x, u) for x, u in zip(d["base"], d["uf"])]
    d = d[d["UF"].isin(UFS)]
    d["data"] = pd.to_datetime(dict(year=d["ano"], month=d["mes"], day=d["dia"]), errors="coerce").dt.date
    d = d[d["data"] < hoje]
    # SLA do mês só com dias de prazo JÁ ENCERRADOS: o "dem_mes" do robô (04h) inclui o prazo de hoje,
    # que ainda nem saiu para entrega e entraria como falha
    dm = d[d["data"] >= dt.date(ano, mes, 1)]
    mtd = dm.groupby("base")[["pacotes", "no_prazo"]].sum()
    b["dem_mes"] = b["base"].map(mtd["pacotes"]).fillna(0)
    b["suc_mes"] = b["base"].map(mtd["no_prazo"]).fillna(0)
    R = b[num].sum()
    sla_mes = _pct(R.suc_mes, R.dem_mes)
    sla_bm = _pct(dm["no_prazo_base"].sum(), dm["pacotes_base"].sum())
    por_dia = d.groupby("data")[["pacotes", "no_prazo"]].sum()
    dias = [x for x in sorted(por_dia.index) if por_dia.loc[x, "pacotes"] > 0]
    sla_dia = lambda x: _pct(por_dia.loc[x, "no_prazo"], por_dia.loc[x, "pacotes"]) if x in por_dia.index else None
    d_sp = (sla_dia(ontem) - sla_dia(sp)) if sla_dia(ontem) is not None and sla_dia(sp) is not None else None
    u7 = dias[-7:]
    j7 = por_dia.loc[u7].sum() if u7 else None
    taxa7 = j7["no_prazo"] / j7["pacotes"] if u7 and j7["pacotes"] else None
    vol7 = j7["pacotes"] / len(u7) if u7 else 0
    import calendar
    restam = max(0, calendar.monthrange(ano, mes)[1] - hoje.day + 1)
    vol = vol7 * restam
    proj = _pct(R.suc_mes + taxa7 * vol, R.dem_mes + vol) if taxa7 is not None else None
    maxp = _pct(R.suc_mes + vol, R.dem_mes + vol) if vol else None

    m = mensal.copy()
    m["UF"] = [_uf(x, u) for x, u in zip(m["base"], m["uf"])]
    ma = m[(m["ano"] == ano_ant) & (m["mes"] == mes_ant) & m["UF"].isin(UFS)]
    sla_ant = _pct(ma["no_prazo"].sum(), ma["pacotes"].sum())

    # backlog
    tot_bl = bl["total"].sum() if not bl.empty else 0
    venc = bl["vencidos"].sum() if not bl.empty else 0
    extr = bl["extravio"].sum() if not bl.empty else 0

    # perdas oficiais
    p = perdas.copy()
    p["UF"] = [_uf(x, u) for x, u in zip(p["base"], p["uf"])]
    p = p[p["UF"].isin(UFS) & (p["ano"] == ano)]
    pm = p.groupby("mes")["valor"].sum()
    fech = [x for x in pm.index if x < mes]
    linha_perdas = ""
    if fech:
        f1 = max(fech)
        linha_perdas = f"*Perdas:* {MES[f1]} {_r(pm[f1])}"
        antes = [x for x in fech if x < f1]
        if antes and pm[max(antes)]:
            dv = 100 * (pm[f1] - pm[max(antes)]) / pm[max(antes)]
            linha_perdas += f" ({'▲' if dv > 0 else '▼'} {abs(dv):.0f}% vs {MES[max(antes)]})"
        if mes in pm.index:
            linha_perdas += f" · {MES[mes]} até agora {_r(pm[mes])}"

    # causas do mês
    c = causas.copy()
    c["UF"] = [_uf(x, u) for x, u in zip(c["base"], c["uf"])]
    c = c[(c["ano"] == ano) & (c["mes"] == mes) & c["UF"].isin(UFS)]
    c["grupo"] = c["causa"].map(lambda k: CAUSA_GRUPO.get(int(k), "Outros"))
    tot_c = c["pacotes"].sum()
    g_c = c.groupby("grupo")["pacotes"].sum().sort_values(ascending=False)
    cb = c.groupby(["base", "grupo"])["pacotes"].sum().reset_index().sort_values("pacotes", ascending=False).drop_duplicates("base").set_index("base")["grupo"]

    L = [f"📊 *PAINEL DA DIRETORIA* · {SEM[hoje.weekday()]} {hoje:%d/%m}"]
    if painel:
        L.append(f"🔗 {painel}")
    L.append("")
    L.append(f"*SLA Anjun ontem ({SEM[ontem.weekday()]} {ontem:%d/%m}):* {_bola(sla_ont)} {_p(sla_ont)}"
             + (f" ({_seta(d_sp)} vs {SEM[sp.weekday()]} da semana passada)" if d_sp is not None else ""))
    ln = f"*SLA do mês:* {_bola(sla_mes)} {_p(sla_mes)}"
    if sla_ant is not None:
        ln += f" · {MES[mes_ant]} fechou {_p(sla_ant)}"
    if proj is not None:
        ln += f" · projeção {_p(proj)}"
    L.append(ln)
    if maxp is not None and maxp < PERIGO:
        L.append(f"⚠️ 95% no mês já não é possível: o máximo é {_p(maxp)}")
    L.append(f"*SLA Base do mês:* {_bola(sla_bm)} {_p(sla_bm)}")
    if tot_bl:
        L.append(f"*Backlog:* {_n(tot_bl)} pacotes · {_p(_pct(venc, tot_bl), 0)} vencido · {_n(extr)} há +7 dias")
    if linha_perdas:
        L.append(linha_perdas)

    L += ["", "*Estados* (SLA mês · ontem)"]
    est = b.groupby("UF")[num].sum()
    for u in sorted(UFS, key=lambda u: _pct(est.loc[u, "suc_mes"], est.loc[u, "dem_mes"]) if u in est.index and est.loc[u, "dem_mes"] else 999):
        if u not in est.index:
            continue
        sm, so = _pct(est.loc[u, "suc_mes"], est.loc[u, "dem_mes"]), _pct(est.loc[u, "suc_ontem"], est.loc[u, "dem_ontem"])
        L.append(f"{_bola(sm)} {u} {_p(sm)} · {_p(so)}")

    # quedas bruscas ontem (base x média dos 7 dias anteriores e x mesmo dia da semana passada)
    db = d.groupby(["base", "data"])[["pacotes", "no_prazo"]].sum()
    ant7 = [x for x in dias if x < ontem][-7:]
    quedas = []
    for base in b["base"]:
        try:
            o = db.loc[(base, ontem)]
        except KeyError:
            continue
        if o["pacotes"] < 150:
            continue
        v = _pct(o["no_prazo"], o["pacotes"])
        h = db.loc[base] if base in db.index.get_level_values(0) else None
        j = h.loc[[x for x in ant7 if x in h.index]].sum() if h is not None else None
        med = _pct(j["no_prazo"], j["pacotes"]) if j is not None else None
        vs = _pct(h.loc[sp, "no_prazo"], h.loc[sp, "pacotes"]) if h is not None and sp in h.index and h.loc[sp, "pacotes"] >= 50 else None
        if v is None or med is None:
            continue
        if med - v >= 15 and (vs is None or vs - v >= 10):
            quedas.append((base, v, med, o["pacotes"] - o["no_prazo"]))
    if quedas:
        L += ["", "🚨 *Despencaram ontem*"]
        for base, v, med, f in sorted(quedas, key=lambda x: -(x[2] - x[1]) * x[3])[:4]:
            L.append(f"• {base}: {_p(v)} (média 7 dias {_p(med)}) · {_n(f)} fora do prazo")

    # onde atacar
    b["falhas"] = b["dem_mes"] - b["suc_mes"]
    tf = b["falhas"].sum()
    top = b[b["dem_mes"] >= 100].sort_values("falhas", ascending=False).head(5)
    # plano de ação (exportado pelo Despertador do painel em _PLANO_ACAO.json)
    hs = hoje.isoformat()
    abertas = [a for a in (plano or []) if a.get("status") not in ("Concluída", "Cancelada")]
    dono = {}
    for a in abertas:
        bb = str(a.get("base", "")).upper()
        if bb and (bb not in dono or str(a.get("prazo", "")) < str(dono[bb].get("prazo", ""))):
            dono[bb] = a

    def sit(a):
        if a.get("prazo") and a["prazo"] < hs:
            return f"⏰ {a.get('resp') or 'sem dono'} (atrasada)"
        return f"✅ {a.get('resp') or 'sem dono'}"

    if tf and len(top):
        L += ["", "🎯 *Onde atacar* (mais atrasos no mês)"]
        for i, (_, x) in enumerate(top.iterrows(), 1):
            L.append(f"{i}. {x['base']} — {_p(_pct(x['suc_mes'], x['dem_mes']))} · {_n(x['falhas'])} atrasos ({100 * x['falhas'] / tf:.0f}%)"
                     + (f" · {cb[x['base']]}" if x["base"] in cb.index else "")
                     + ("" if plano is None else (f" · {sit(dono[x['base']])}" if x["base"] in dono else " · ❗ sem ação")))
    if plano:
        atr = sorted([a for a in abertas if a.get("prazo") and a["prazo"] < hs], key=lambda a: a["prazo"])
        L += ["", f"📋 *Plano de ação:* {len(abertas)} abertas · {len(atr)} atrasada{'s' if len(atr) != 1 else ''}"]
        for a in atr[:3]:
            pz = a["prazo"][8:10] + "/" + a["prazo"][5:7]
            L.append(f"⏰ {a.get('base')} — {a.get('resp') or 'sem dono'} · prazo {pz}")
    if tot_c:
        g1 = g_c.index[0]
        c1 = c.groupby("causa")["pacotes"].sum().idxmax()
        L += ["", f"🧠 {g1} causa {100 * g_c.iloc[0] / tot_c:.0f}% das falhas do mês. Maior causa: {CAUSA_NOME.get(int(c1), '-')}."]
    L += ["", f"_SLA: {X.get('sla_ref', '')} · Backlog: {X.get('bl_ref', '')}_"]
    return "\n".join(L)
