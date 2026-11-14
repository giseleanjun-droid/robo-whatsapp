# -*- coding: utf-8 -*-
"""
matriz.py - monta a MATRIZ única (SLA + BACKLOG de todos os estados) em .xlsx, já formatada.
O robô envia esse .xlsx para o Drive convertendo em Planilha Google (mesmo ID sempre).

Uso local (teste):  python matriz.py AC [GO ...]   -> _saida/MATRIZ_SLA_BACKLOG.xlsx
"""
import sys
import datetime as dt
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

import motor
import mensagem

# cores (as mesmas dos scripts antigos)
ESCURO, BRANCO = "202124", "FFFFFF"
VERDE, AMARELO, VERMELHO = "D9EAD3", "FFF2CC", "F4CCCC"
CINZA = "EFEFEF"
COR_EMB = {"TIKTOK": "000000", "SHEIN": "E8336D"}
F_PCT, F_INT, F_RS = "0.0%", "#,##0", '"R$" #,##0.00'
F_DATA_H, F_DATA = "dd/mm/yyyy hh:mm", "dd/mm/yyyy"
FILL = lambda c: PatternFill("solid", start_color=c, end_color=c)
FINO = Side(style="thin", color="D0D0D0")

COLS_HIST = ["data", "uf", "base", "sla_a_ontem", "sla_b_ontem", "sla_a_mes", "sla_b_mes", "atrasado", "hoje",
             "bl_total", "bl_valor", "bl_extravio", "bl_vencidos", "bl_perdas", "bl_valor_perdas", "tiktok_aberto"]


# ------------------------------------------------------------------ helpers de escrita
def _cabecalho(ws, linha, cols, cor=ESCURO):
    for j, c in enumerate(cols, 1):
        cel = ws.cell(linha, j, c)
        cel.fill, cel.font = FILL(cor), Font(bold=True, color=BRANCO)
        cel.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[linha].height = 32


def _escrever(ws, df, linha_cab=2, cor=ESCURO, formatos=None, larguras=None, filtro=True, congelar=True):
    """Escreve df a partir de linha_cab (cabeçalho) e aplica formatos por nome de coluna."""
    cols = list(df.columns)
    _cabecalho(ws, linha_cab, cols, cor)
    for i, row in enumerate(df.itertuples(index=False), linha_cab + 1):
        for j, v in enumerate(row, 1):
            if v is None or (isinstance(v, float) and np.isnan(v)) or v is pd.NaT or (v is pd.NA):
                continue
            if isinstance(v, pd.Timestamp):
                v = v.to_pydatetime()
            elif isinstance(v, (np.integer,)):
                v = int(v)
            elif isinstance(v, (np.floating,)):
                v = float(v)
            elif isinstance(v, (np.bool_,)):
                v = bool(v)
            ws.cell(i, j, v)
    ult = linha_cab + len(df)
    for nome, fmt in (formatos or {}).items():
        if nome in cols:
            j = cols.index(nome) + 1
            for i in range(linha_cab + 1, ult + 1):
                ws.cell(i, j).number_format = fmt
    for j, c in enumerate(cols, 1):
        ws.column_dimensions[get_column_letter(j)].width = (larguras or {}).get(c, max(11, min(28, len(str(c)) + 3)))
    if filtro and len(df):
        ws.auto_filter.ref = f"A{linha_cab}:{get_column_letter(len(cols))}{ult}"
    if congelar:
        ws.freeze_panes = ws.cell(linha_cab + 1, 1)
    return ult


def _semaforo(ws, rng):
    """Verde >= 98% · amarelo 80-98% · vermelho < 80% (igual aplicarSemaforo)."""
    ws.conditional_formatting.add(rng, CellIsRule(operator="greaterThanOrEqual", formula=["0.98"], fill=FILL(VERDE)))
    ws.conditional_formatting.add(rng, CellIsRule(operator="between", formula=["0.8", "0.97999"], fill=FILL(AMARELO)))
    ws.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["0.8"], fill=FILL(VERMELHO)))


def _semaforo_cols(ws, df, nomes, linha_cab=2):
    for n in nomes:
        if n in df.columns and len(df):
            L = get_column_letter(list(df.columns).index(n) + 1)
            _semaforo(ws, f"{L}{linha_cab + 1}:{L}{linha_cab + len(df)}")


def _pinta_coluna(ws, df, nome, cor, fonte=None, linha_cab=2):
    if nome not in df.columns:
        return
    j = list(df.columns).index(nome) + 1
    for i in range(linha_cab + 1, linha_cab + len(df) + 1):
        c = ws.cell(i, j)
        c.fill = FILL(cor)
        if fonte:
            c.font = Font(color=fonte, bold=True)


def _titulo(ws, texto, sub=None):
    ws.cell(1, 1, texto).font = Font(bold=True, size=12)
    if sub:
        ws.cell(1, 6, sub).font = Font(bold=True, color="CC0000", size=9)
    ws.sheet_view.showGridLines = False


def _texto_regras(ws, rng, regras):
    """Pinta a célula se o texto contém a palavra: [(palavra, fundo, fonte)]."""
    col1 = rng.split(":")[0]
    letra = "".join(ch for ch in col1 if ch.isalpha())
    lin = "".join(ch for ch in col1 if ch.isdigit())
    for palavra, fundo, fonte in regras:
        ws.conditional_formatting.add(rng, FormulaRule(
            formula=[f'ISNUMBER(SEARCH("{palavra}",{letra}{lin}))'], fill=FILL(fundo),
            font=Font(color=fonte or "000000", bold=True)))


# ------------------------------------------------------------------ abas
def aba_resumo_estados(wb, S, B, carimbo):
    ws = wb.active
    ws.title = "📊 RESUMO ESTADOS"
    _titulo(ws, "📊 RESUMO POR ESTADO — SLA + BACKLOG", carimbo)
    sb, bb = S["bases"], B["bases"]
    s = sb.groupby("uf")[["dem_mes", "suc_mes", "demb_mes", "sucb_mes", "dem_ontem", "suc_ontem", "demb_ontem",
                          "sucb_ontem", "atrasado", "hoje", "backlog"]].sum()
    b = bb.groupby("uf")[["total", "valor", "vencidos", "vence_hoje", "extravio", "perdas", "valor_perdas", "tiktok"]].sum()
    df = s.join(b, how="outer").fillna(0).reset_index()
    tk = S["tiktok"].groupby("uf").size()
    d = lambda a, c: np.where(df[c] > 0, df[a] / df[c].replace(0, 1), np.nan)
    out = pd.DataFrame({
        "UF": df["uf"], "Nº BASES": df["uf"].map(sb.groupby("uf").size()).fillna(0).astype(int),
        "SLA ANJUN ONTEM": d("suc_ontem", "dem_ontem"), "SLA BASE ONTEM": d("sucb_ontem", "demb_ontem"),
        "SLA ANJUN MÊS": d("suc_mes", "dem_mes"), "SLA BASE MÊS": d("sucb_mes", "demb_mes"),
        "PACOTES SLA (MÊS)": df["dem_mes"], "🔴 ATRASADOS": df["atrasado"], "🟡 VENCE HOJE": df["hoje"],
        "📦 BACKLOG": df["total"], "VALOR EM RISCO": df["valor"], "🔥 VENCIDOS": df["vencidos"],
        "🚨 +7 DIAS": df["extravio"], "💸 PERDAS (QTD)": df["perdas"], "💸 PERDAS (R$)": df["valor_perdas"],
        "🎵 TIKTOK EM ABERTO": df["uf"].map(tk).fillna(0).astype(int),
    })
    tot = {c: out[c].sum() for c in out.columns if c not in ("UF",) and not c.startswith("SLA")}
    tot["UF"] = "TOTAL"
    for nome, (a, c) in {"SLA ANJUN ONTEM": ("suc_ontem", "dem_ontem"), "SLA BASE ONTEM": ("sucb_ontem", "demb_ontem"),
                         "SLA ANJUN MÊS": ("suc_mes", "dem_mes"), "SLA BASE MÊS": ("sucb_mes", "demb_mes")}.items():
        tot[nome] = df[a].sum() / df[c].sum() if df[c].sum() else np.nan
    out = pd.concat([out, pd.DataFrame([tot])], ignore_index=True)
    ult = _escrever(ws, out, formatos={**{c: F_PCT for c in out.columns if c.startswith("SLA")},
                                       "VALOR EM RISCO": F_RS, "💸 PERDAS (R$)": F_RS}, filtro=False)
    _semaforo_cols(ws, out, [c for c in out.columns if c.startswith("SLA")])
    for j in range(1, len(out.columns) + 1):
        ws.cell(ult, j).font = Font(bold=True)
        ws.cell(ult, j).fill = FILL(CINZA)
    return out


def aba_sla_bases(wb, S, carimbo):
    ws = wb.create_sheet("📈 SLA POR BASE")
    _titulo(ws, "📈 SLA POR BASE — Gestão do dia e acumulado do mês (Meta 98%)", carimbo)
    b = S["bases"].copy()
    eb = S["emb_base"].set_index(["base", "emb"])
    def e(base, emb, k):
        try:
            return eb.loc[(base, emb), k]
        except KeyError:
            return np.nan
    out = pd.DataFrame({
        "UF": b["uf"], "BASE": b["base"],
        "SLA ANJUN ONTEM": b["sla_a_ontem"], "SLA BASE ONTEM": b["sla_b_ontem"],
        "SLA ANJUN MÊS": b["sla_a_mes"], "SLA BASE MÊS": b["sla_b_mes"],
        "PACOTES DO DIA (ANJUN)": b["dem_hoje"], "✅ SUCESSO DIA": b["suc_hoje"], "❌ FALHA DIA": b["dem_hoje"] - b["suc_hoje"],
        "PACOTES MÊS (ANJUN)": b["dem_mes"], "✅ SUCESSO MÊS": b["suc_mes"], "❌ FALHA MÊS": b["dem_mes"] - b["suc_mes"],
        "📦 BACKLOG": b["backlog"], "🔴 ATRASADO": b["atrasado"], "🟡 HOJE": b["hoje"], "🟢 AMANHÃ": b["amanha"],
        "🔵 SEMANA": b["semana"], "🏢 NA BASE": b["na_base"], "🏍️ EM ROTA": b["em_rota"], "🚨 +7 DIAS": b["extravio"],
        "CHEGOU VENCIDO (MÊS)": b["cv_mes"],
        "🎵 TIKTOK SLA ANJUN MÊS": [e(x, "TIKTOK", "sla_a_mes") for x in b["base"]],
        "🎵 TIKTOK SLA BASE MÊS": [e(x, "TIKTOK", "sla_b_mes") for x in b["base"]],
        "🎵 TIKTOK ATRASADOS": [e(x, "TIKTOK", "atrasado") for x in b["base"]],
        "🛍️ SHEIN SLA ANJUN MÊS": [e(x, "SHEIN", "sla_a_mes") for x in b["base"]],
        "🛍️ SHEIN SLA BASE MÊS": [e(x, "SHEIN", "sla_b_mes") for x in b["base"]],
    }).sort_values(["UF", "SLA ANJUN MÊS"], ascending=[True, True])
    pcts = [c for c in out.columns if "SLA" in c]
    _escrever(ws, out, formatos={c: F_PCT for c in pcts}, larguras={"BASE": 14})
    _semaforo_cols(ws, out, pcts)


def aba_sla_diario(wb, S, carimbo):
    ws = wb.create_sheet("📅 SLA DIA A DIA")
    hoje = S["hoje"]
    _titulo(ws, f"📅 DIA A DIA — performance SLA {hoje:%m/%Y} (Meta 98%)", carimbo)
    d = S["diario"].copy()
    if d.empty:
        return
    dias = list(range(1, hoje.day + 1))
    linhas = []
    nomes = [("SLA Anjun %", "suc", "dem"), ("SLA Base %", "sucb", "demb"),
             ("🎵 TikTok SLA Anjun %", "suc_TIKTOK", "dem_TIKTOK"), ("🎵 TikTok SLA Base %", "sucb_TIKTOK", "demb_TIKTOK")]
    cont = [("Pacotes do dia SLA Anjun", "dem"), ("Entregue (Baixas)", "entregue"), ("Recebido", "recebido"),
            ("Recebido após 09h (+1 dia SLA Base)", "recebido_apos9"),
            ("Chegou atrasado do dia", "chegou_atrasado"), ("Pacotes com problema", "problema")]
    for c in {x for _, a, b in nomes for x in (a, b)} | {b for _, b in cont}:
        if c not in d.columns:
            d[c] = 0
    uf = S["bases"].set_index("base")["uf"]
    for base, g in d.groupby("base"):
        g = g.set_index("dia")
        for rot, a, b in nomes:
            vals = [(g.at[x, a] / g.at[x, b]) if (x in g.index and g.at[x, b] > 0) else None for x in dias]
            tot = g[a].sum() / g[b].sum() if g[b].sum() else None
            linhas.append([uf.get(base, base[:2]), base, rot, tot] + vals)
        for rot, c in cont:
            vals = [int(g.at[x, c]) if x in g.index and g.at[x, c] else None for x in dias]
            linhas.append([uf.get(base, base[:2]), base, rot, int(g[c].sum())] + vals)
    cols = ["UF", "BASE", "INDICADOR", "ACUMULADO"] + [f"{x:02d}/{hoje.month:02d}" for x in dias]
    out = pd.DataFrame(linhas, columns=cols)
    _escrever(ws, out, larguras={"INDICADOR": 26, "BASE": 13, **{c: 7 for c in cols[4:]}})
    n = len(out)
    pct_rows = [i + 3 for i, r in enumerate(linhas) if "%" in r[2]]
    for i in range(3, n + 3):
        fmt = F_PCT if i in pct_rows else F_INT
        for j in range(4, len(cols) + 1):
            ws.cell(i, j).number_format = "0%" if fmt == F_PCT and j > 4 else fmt
    ult = get_column_letter(len(cols))
    for i in pct_rows:
        _semaforo(ws, f"D{i}:{ult}{i}")
    for i, r in enumerate(linhas, 3):
        if r[2] == "Pacotes do dia SLA Anjun":
            pass
        if r[2] == "SLA Anjun %":
            for j in range(1, 4):
                ws.cell(i, j).font = Font(bold=True)


def aba_entradas(wb, S, carimbo):
    """Quanto chega na base DEPOIS das 09h (ganha +1 dia no SLA Base) — para cobrar transporte e HUB."""
    e = S.get("entradas")
    if e is None or e.empty:
        return
    hoje = S["hoje"]
    ontem = (hoje - dt.timedelta(days=1)).day if hoje.day > 1 else None
    corte = f"{motor.HORA_CORTE:02d}h"
    ws = wb.create_sheet(f"🚚 ENTRADA APÓS {corte}")
    _titulo(ws, f"🚚 PACOTES QUE ENTRARAM NA BASE A PARTIR DAS {corte} (ganham +1 dia no SLA Base) — por HUB e base", carimbo)
    chave = ["uf", "hub", "base"]
    mes = e.groupby(chave)[["recebidos", "apos_corte"]].sum()
    on = e[e["dia"] == ontem].set_index(chave) if ontem else e.iloc[0:0].set_index(chave)
    df = mes.join(on[["recebidos", "apos_corte", "primeira", "ultima", "media"]], rsuffix="_ontem").reset_index()
    h = lambda s: s / 24   # hora decimal -> fração do dia (formato hh:mm)
    out = pd.DataFrame({
        "UF": df["uf"], "HUB DE TRÂNSITO": df["hub"], "BASE": df["base"],
        "ONTEM: RECEBIDOS": df["recebidos_ontem"], f"ONTEM: APÓS {corte}": df["apos_corte_ontem"],
        f"ONTEM: % APÓS {corte}": np.where(df["recebidos_ontem"] > 0, df["apos_corte_ontem"] / df["recebidos_ontem"].replace(0, 1), np.nan),
        "ONTEM: 1ª ENTRADA": h(df["primeira"]), "ONTEM: ÚLTIMA ENTRADA": h(df["ultima"]), "ONTEM: HORA MÉDIA": h(df["media"]),
        "MÊS: RECEBIDOS": df["recebidos"], f"MÊS: APÓS {corte}": df["apos_corte"],
        f"MÊS: % APÓS {corte}": np.where(df["recebidos"] > 0, df["apos_corte"] / df["recebidos"].replace(0, 1), np.nan),
    }).sort_values(["UF", f"MÊS: APÓS {corte}"], ascending=[True, False])
    _escrever(ws, out, cor="B45F06", larguras={"HUB DE TRÂNSITO": 16, "BASE": 13},
              formatos={f"ONTEM: % APÓS {corte}": F_PCT, f"MÊS: % APÓS {corte}": F_PCT,
                        "ONTEM: 1ª ENTRADA": "hh:mm", "ONTEM: ÚLTIMA ENTRADA": "hh:mm", "ONTEM: HORA MÉDIA": "hh:mm"})
    for nome in (f"ONTEM: % APÓS {corte}", f"MÊS: % APÓS {corte}"):   # aqui quanto MAIOR, pior
        L = get_column_letter(list(out.columns).index(nome) + 1)
        rng = f"{L}3:{L}{len(out) + 2}"
        ws.conditional_formatting.add(rng, CellIsRule(operator="greaterThanOrEqual", formula=["0.5"], fill=FILL(VERMELHO)))
        ws.conditional_formatting.add(rng, CellIsRule(operator="between", formula=["0.2", "0.49999"], fill=FILL(AMARELO)))
        ws.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["0.2"], fill=FILL(VERDE)))

    # resumo por HUB (ao lado) — para a conversa com transporte/HUB
    i0 = len(out.columns) + 2
    hb = e.groupby(["uf", "hub"])[["recebidos", "apos_corte"]].sum().reset_index()
    hb["pct"] = np.where(hb["recebidos"] > 0, hb["apos_corte"] / hb["recebidos"].replace(0, 1), np.nan)
    hb = hb.sort_values("apos_corte", ascending=False)
    for j, c in enumerate(["UF ", "HUB ", "RECEBIDOS (MÊS)", f"APÓS {corte} (MÊS)", f"% APÓS {corte}"], i0):
        cel = ws.cell(2, j, c)
        cel.fill, cel.font = FILL("B45F06"), Font(bold=True, color=BRANCO)
        cel.alignment = Alignment(horizontal="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(j)].width = 14
    for i, r in enumerate(hb.itertuples(index=False), 3):
        for j, v in enumerate([r.uf, r.hub, int(r.recebidos), int(r.apos_corte), float(r.pct)], i0):
            ws.cell(i, j, v)
        ws.cell(i, i0 + 4).number_format = F_PCT

    # dia a dia: % que entrou após o corte, por base
    ws2 = wb.create_sheet(f"🚚 ENTRADA DIA A DIA")
    _titulo(ws2, f"🚚 % DOS PACOTES QUE ENTRARAM A PARTIR DAS {corte}, DIA A DIA (e a hora da 1ª entrada)", carimbo)
    g = e.groupby(["uf", "base", "dia"]).agg(rec=("recebidos", "sum"), apos=("apos_corte", "sum"), prim=("primeira", "min")).reset_index()
    dias = list(range(1, hoje.day))
    linhas = []
    for (uf, base), x in g.groupby(["uf", "base"]):
        x = x.set_index("dia")
        tot = x["apos"].sum() / x["rec"].sum() if x["rec"].sum() else None
        linhas.append([uf, base, f"% após {corte}", tot] +
                      [(x.at[d, "apos"] / x.at[d, "rec"]) if d in x.index and x.at[d, "rec"] else None for d in dias])
        linhas.append([uf, base, "1ª entrada (hora)", None] + [(x.at[d, "prim"] / 24) if d in x.index else None for d in dias])
    cols = ["UF", "BASE", "INDICADOR", "MÊS"] + [f"{d:02d}/{hoje.month:02d}" for d in dias]
    o2 = pd.DataFrame(linhas, columns=cols)
    _escrever(ws2, o2, cor="B45F06", larguras={"INDICADOR": 18, **{c: 7 for c in cols[4:]}})
    for i, r in enumerate(linhas, 3):
        fmt = "0%" if r[2].startswith("%") else "hh:mm"
        for j in range(4, len(cols) + 1):
            ws2.cell(i, j).number_format = fmt


def aba_embarcadores(wb, S, carimbo):
    ws = wb.create_sheet("🛍️ EMBARCADORES")
    _titulo(ws, "🛍️ SLA POR EMBARCADOR (mês)", carimbo)
    e = S["emb_uf"].copy()
    tot_uf = e.groupby("uf")["dem_mes"].transform("sum")
    tot_falha_uf = e.groupby("uf")["dem_mes"].transform("sum")
    out = pd.DataFrame({
        "UF": e["uf"], "EMBARCADOR": e["emb"], "PACOTES SLA (MÊS)": e["dem_mes"],
        "% DO VOLUME": np.where(tot_uf > 0, e["dem_mes"] / tot_uf.replace(0, 1), np.nan),
        "SLA ANJUN MÊS": e["sla_a_mes"], "SLA ANJUN ONTEM": e["sla_a_ontem"],
        "SLA BASE MÊS": e["sla_b_mes"], "SLA BASE ONTEM": e["sla_b_ontem"],
        "FALHAS (MÊS)": e["dem_mes"] - e["suc_mes"],
        "PONTOS PERDIDOS NO SLA": np.where(tot_falha_uf > 0, (e["dem_mes"] - e["suc_mes"]) / tot_falha_uf.replace(0, 1), np.nan),
        "CHEGOU VENCIDO": e["cv_mes"], "BACKLOG": e["backlog"], "ATRASADOS": e["atrasado"], "VENCE HOJE": e["hoje"],
    })
    out = out[out["PACOTES SLA (MÊS)"] + out["BACKLOG"] > 0].sort_values(["UF", "PACOTES SLA (MÊS)"], ascending=[True, False])
    pcts = ["% DO VOLUME", "SLA ANJUN MÊS", "SLA ANJUN ONTEM", "SLA BASE MÊS", "SLA BASE ONTEM", "PONTOS PERDIDOS NO SLA"]
    _escrever(ws, out, formatos={c: F_PCT for c in pcts}, larguras={"EMBARCADOR": 22})
    _semaforo_cols(ws, out, pcts[1:5])


def aba_ranking(wb, S, carimbo):
    ws = wb.create_sheet("🏆 RANKING E IMPACTO")
    _titulo(ws, "🏆 RANKING SLA ONTEM e 🚨 BASES QUE DERRUBAM O ESTADO", carimbo)
    b = S["bases"].copy()
    tot_uf = b.groupby("uf")["dem_mes"].transform("sum")
    b["impacto"] = np.where(tot_uf > 0, (b["dem_mes"] - b["suc_mes"]) / tot_uf.replace(0, 1), 0)
    r = b[b["dem_ontem"] > 0].sort_values(["uf", "sla_a_ontem"], ascending=[True, False])
    rk = pd.DataFrame({"UF": r["uf"], "BASE": r["base"], "SLA ANJUN ONTEM": r["sla_a_ontem"], "PACOTES ONTEM": r["dem_ontem"],
                       "SLA BASE ONTEM": r["sla_b_ontem"], "PACOTES BASE ONTEM": r["demb_ontem"]})
    rk.insert(1, "#", rk.groupby("UF").cumcount() + 1)
    ult = _escrever(ws, rk, formatos={"SLA ANJUN ONTEM": F_PCT, "SLA BASE ONTEM": F_PCT}, filtro=True)
    _semaforo_cols(ws, rk, ["SLA ANJUN ONTEM", "SLA BASE ONTEM"])
    # impacto ao lado
    i0 = len(rk.columns) + 2
    imp = b[b["dem_mes"] > 0].sort_values(["uf", "impacto"], ascending=[True, False])
    imp = pd.DataFrame({"UF ": imp["uf"], "BASE ": imp["base"], "PACOTES (MÊS)": imp["dem_mes"],
                        "FALHAS (MÊS)": imp["dem_mes"] - imp["suc_mes"], "SLA DA BASE": imp["sla_a_mes"],
                        "IMPACTO NO ESTADO (pp)": imp["impacto"]})
    for j, c in enumerate(imp.columns, i0):
        cel = ws.cell(2, j, c)
        cel.fill, cel.font = FILL("CC0000"), Font(bold=True, color=BRANCO)
        cel.alignment = Alignment(horizontal="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(j)].width = 14
    for i, row in enumerate(imp.itertuples(index=False), 3):
        for j, v in enumerate(row, i0):
            ws.cell(i, j, v.item() if hasattr(v, "item") else v)
        ws.cell(i, i0 + 4).number_format = F_PCT
        ws.cell(i, i0 + 5).number_format = "0.00%"
    L = get_column_letter(i0 + 4)
    if len(imp):
        _semaforo(ws, f"{L}3:{L}{len(imp) + 2}")


def aba_cidades(wb, S, carimbo):
    ws = wb.create_sheet("🏙️ CIDADES E MOTORISTAS")
    _titulo(ws, "🏙️ CIDADES e 🚚 MOTORISTAS OFENSORES (mês)", carimbo)
    c = S["cidades"]
    c = c[(c["entregues"] > 0) | (c["backlog"] > 0)].copy()
    c["fails"] = c["fora_a"] + c["fora_b"]
    c = c.sort_values("fails", ascending=False)
    out = pd.DataFrame({"UF": c["uf"], "BASE": c["base"], "CIDADE": c["cidade"], "ENTREGUES": c["entregues"],
                        "SLA ANJUN %": c["sla_a"], "SLA BASE %": c["sla_b"], "BACKLOG": c["backlog"],
                        "ATRASADOS": c["atrasado"], "CHEGOU VENCIDO": c["chegou_vencido"], "FAILS TOTAIS": c["fails"],
                        "🎵 TIKTOK ENTREGUES": c["TIKTOK_entregues"],
                        "🎵 TIKTOK SLA ANJUN %": np.where(c["TIKTOK_entregues"] > 0, c["TIKTOK_ok"] / c["TIKTOK_entregues"].replace(0, 1), np.nan),
                        "🎵 TIKTOK BACKLOG": c["TIKTOK_backlog"]})
    _escrever(ws, out, formatos={"SLA ANJUN %": F_PCT, "SLA BASE %": F_PCT, "🎵 TIKTOK SLA ANJUN %": F_PCT},
              larguras={"CIDADE": 24})
    _semaforo_cols(ws, out, ["SLA ANJUN %", "SLA BASE %", "🎵 TIKTOK SLA ANJUN %"])

    ws2 = wb.create_sheet("🚚 MOTORISTAS")
    _titulo(ws2, "🚚 MOTORISTAS OFENSORES EM SLA BASE (volume)", carimbo)
    m = S["motoristas"]
    m = m[(m["entregues"] > 0) | (m["backlog"] > 0)].sort_values("fora_b", ascending=False)
    om = pd.DataFrame({"UF": m["uf"], "BASE": m["base"], "MOTORISTA": m["motorista"], "VALOR ABERTO": m["valor_aberto"],
                       "ENTREGUES": m["entregues"], "SLA ANJUN %": m["sla_a"], "SLA BASE %": m["sla_b"],
                       "FORA SLA BASE": m["fora_b"], "BACKLOG": m["backlog"], "ATRASADOS": m["atrasado"],
                       "🎵 TIKTOK ENTREGUES": m["TIKTOK_entregues"], "🎵 TIKTOK BACKLOG": m["TIKTOK_backlog"]})
    _escrever(ws2, om, formatos={"SLA ANJUN %": F_PCT, "SLA BASE %": F_PCT, "VALOR ABERTO": F_RS}, larguras={"MOTORISTA": 26})
    _semaforo_cols(ws2, om, ["SLA ANJUN %", "SLA BASE %"])


def aba_backlog(wb, B, carimbo):
    ws = wb.create_sheet("📦 BACKLOG POR BASE")
    _titulo(ws, "📦 ANÁLISE DO BACKLOG POR BASE", carimbo)
    b = B["bases"].sort_values(["uf", "total"], ascending=[True, False])
    out = pd.DataFrame({
        "UF": b["uf"], "BASE": b["base"], "TOTAL BACKLOG": b["total"], "VALOR (R$)": b["valor"], "OFENSOR": b["ofensor"],
        "🔴 > 7 Dias": b["age_mais7"], "🟠 7 Dias": b["age7"], "🟠 6 Dias": b["age6"], "🟡 5 Dias": b["age5"],
        "🟡 4 Dias": b["age4"], "🟡 3 Dias": b["age3"], "🟡 2 Dias": b["age2"], "🟢 1 Dia": b["age1"],
        "🔥 VENCIDOS": b["vencidos"], "⚠️ VENCE HOJE": b["vence_hoje"], "🏴‍☠️ EXTRAVIO (>7d)": b["extravio"],
        "🏍️ EM ROTA": b["rota"], "📦 ESTOQUE": b["estoque"], "🌍 OUTRA REGIÃO": b["outra_regiao"],
        "💸 PERDA (QTD)": b["perdas"], "💸 PERDA CONFIRMADA (R$)": b["valor_perdas"],
        "📍 D0": b["d0"], "📍 D+1": b["d1"], "📍 D+2": b["d2"], "📍 D+3 ou mais": b["d3"], "🎵 TIKTOK": b["tiktok"],
    })
    _escrever(ws, out, formatos={"VALOR (R$)": F_RS, "💸 PERDA CONFIRMADA (R$)": F_RS}, larguras={"BASE": 13})
    for nome, cor, fonte in [("🔴 > 7 Dias", "CC0000", BRANCO), ("🟠 7 Dias", "FF9900", None), ("🟠 6 Dias", "FF9900", None),
                             ("🟡 5 Dias", "F1C232", None), ("🟡 4 Dias", "F1C232", None), ("🟡 3 Dias", "F1C232", None),
                             ("🟡 2 Dias", "F1C232", None), ("🟢 1 Dia", "34A853", BRANCO), ("🔥 VENCIDOS", "CC0000", BRANCO),
                             ("⚠️ VENCE HOJE", "FF9900", None), ("🏴‍☠️ EXTRAVIO (>7d)", "000000", BRANCO)]:
        _pinta_coluna(ws, out, nome, cor, fonte)

    ws2 = wb.create_sheet("🏙️ PRAZOS CIDADES")
    _titulo(ws2, "🏙️ PRAZO (SLA BASE) MAIS COMUM POR CIDADE — pacotes do backlog", carimbo)
    pc = B["prazos_cidades"].sort_values(["uf", "base", "qtd"], ascending=[True, True, False])
    lbl = pc["prazo"].map(lambda d: "D0" if d == 0 else ("D+1" if d == 1 else ("D+2" if d == 2 else "D+3")))
    _escrever(ws2, pd.DataFrame({"UF": pc["uf"], "BASE": pc["base"], "CIDADE DE DESTINO": pc["cidade"],
                                 "PRAZO (SLA)": lbl, "QTD PACOTES": pc["qtd"]}), cor="674EA7", larguras={"CIDADE DE DESTINO": 28})


def aba_perdas(wb, B, carimbo):
    ws = wb.create_sheet("💸 PERDAS")
    l = B["lista"]
    l = l[l["ALERTA"] == motor.AL_PERDA]
    _titulo(ws, f"💸 PERDAS CONFIRMADAS — {len(l)} pacotes · {mensagem.rs(l['Valor'].sum())}", carimbo)
    out = l[["UF", "BASE", "Waybill", "Embarcador", "Motorista", "Cidade", "Dias Parado", "Valor", "Motivo", "Registrante"]]
    _escrever(ws, out, cor="000000", formatos={"Valor": F_RS}, larguras={"Motivo": 40, "Waybill": 20})


def aba_tiktok(wb, S, carimbo):
    ws = wb.create_sheet("🎵 TIKTOK EM ABERTO")
    tk = S["tiktok"]
    _titulo(ws, f"🎵 TIKTOK EM ABERTO (na base e em rota) — {len(tk)} pacotes, mais urgente primeiro", carimbo)
    out = pd.DataFrame({"PRAZO ANJUN": tk["classe"], "PRAZO BASE": tk["classe_b"], "UF": tk["uf"], "BASE": tk["base"],
                        "Waybill": tk["waybill"], "Status / Local": tk["status"],
                        "ONDE": np.where(tk["em_rota"], "🏍️ EM ROTA", "🏢 NA BASE"), "Motorista": tk["motorista"],
                        "Cidade": tk["cidade"], "Dias Parado": tk["dias"], "Valor": tk["valor"],
                        "Data Anjun": tk["prazo"], "Data Base": tk["prazo_b"]})
    _escrever(ws, out, cor="000000", formatos={"Valor": F_RS, "Data Anjun": F_DATA_H, "Data Base": F_DATA_H},
              larguras={"Waybill": 20, "Motorista": 24, "Cidade": 22})
    _pinta_coluna(ws, out, "PRAZO BASE", "FFF4E5")
    n = len(out) + 2
    if len(out):
        _texto_regras(ws, f"A3:B{n}", [("ATRASADO", "CC0000", BRANCO), ("HOJE", "FF9900", None),
                                       ("AMANH", "D9EAD3", None), ("SEMANA", "CFE2F3", None)])


def aba_whatsapp(wb, msgs, carimbo):
    ws = wb.create_sheet("📲 WHATSAPP")
    _titulo(ws, "📲 MENSAGENS DO DIA — o robô envia às 09h no grupo de cada base", carimbo)
    out = pd.DataFrame(msgs, columns=["UF", "BASE", "PLANILHA DA BASE", "GRUPO WHATSAPP", "LINK DO GRUPO", "MENSAGEM (COPIAR)",
                                      "SLA ANJUN ONTEM", "SLA BASE ONTEM", "BACKLOG", "PERDAS (R$)"])
    _escrever(ws, out, cor="25D366", formatos={"SLA ANJUN ONTEM": F_PCT, "SLA BASE ONTEM": F_PCT, "PERDAS (R$)": F_RS},
              larguras={"MENSAGEM (COPIAR)": 70, "GRUPO WHATSAPP": 24, "LINK DO GRUPO": 22, "PLANILHA DA BASE": 30})
    jp = list(out.columns).index("PLANILHA DA BASE") + 1
    for i, url in enumerate(out["PLANILHA DA BASE"], 3):
        if str(url).startswith("http"):
            c = ws.cell(i, jp)
            c.value = f"📄 abrir {out.iloc[i - 3]['BASE']}"
            c.hyperlink = url
            c.font = Font(color="1155CC", underline="single")
    j = list(out.columns).index("MENSAGEM (COPIAR)") + 1
    for i in range(3, len(out) + 3):
        ws.cell(i, j).alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[i].height = 150
    _semaforo_cols(ws, out, ["SLA ANJUN ONTEM", "SLA BASE ONTEM"])


def aba_historico(wb, hist):
    ws = wb.create_sheet("_HISTORICO")
    _titulo(ws, "_HISTORICO — foto diária de cada base (usada no 'melhorou / piorou'). Não editar.")
    _escrever(ws, hist[COLS_HIST], formatos={"sla_a_ontem": F_PCT, "sla_b_ontem": F_PCT, "sla_a_mes": F_PCT,
                                             "sla_b_mes": F_PCT, "bl_valor": F_RS, "bl_valor_perdas": F_RS})


# ------------------------------------------------------------------ mensagens + histórico
def foto_do_dia(S, B):
    hoje = (S or B)["hoje"].date().isoformat()
    s = S["bases"].set_index("base") if S else pd.DataFrame()
    b = B["bases"].set_index("base") if B else pd.DataFrame()
    tk = S["tiktok"].groupby("base").size() if S else pd.Series(dtype=int)
    bases = sorted(set(s.index) | set(b.index))
    linhas = []
    for x in bases:
        gs = lambda k: s.at[x, k] if x in s.index else np.nan
        gb = lambda k: b.at[x, k] if x in b.index else np.nan
        linhas.append({"data": hoje, "uf": x[:2], "base": x, "sla_a_ontem": gs("sla_a_ontem"), "sla_b_ontem": gs("sla_b_ontem"),
                       "sla_a_mes": gs("sla_a_mes"), "sla_b_mes": gs("sla_b_mes"), "atrasado": gs("atrasado"),
                       "hoje": gs("hoje"), "bl_total": gb("total"), "bl_valor": gb("valor"), "bl_extravio": gb("extravio"),
                       "bl_vencidos": gb("vencidos"), "bl_perdas": gb("perdas"), "bl_valor_perdas": gb("valor_perdas"),
                       "tiktok_aberto": int(tk.get(x, 0))})
    return pd.DataFrame(linhas, columns=COLS_HIST)


def juntar_historico(antigo, foto, dias=62):
    """Troca a foto de hoje (se já existir) e guarda ~2 meses."""
    if antigo is None or antigo.empty:
        return foto
    antigo = antigo[antigo["data"].astype(str) != foto["data"].iloc[0]]
    h = pd.concat([antigo, foto], ignore_index=True)
    corte = (pd.Timestamp(foto["data"].iloc[0]) - pd.Timedelta(days=dias)).date().isoformat()
    return h[h["data"].astype(str) >= corte]


def montar_mensagens(S, B, hist, grupos=None, links=None):
    hoje = (S or B)["hoje"]
    ontem = hoje - dt.timedelta(days=1)
    data_txt, ontem_txt = f"{hoje:%d/%m}", f"{ontem:%d/%m}"
    hojeiso = hoje.date().isoformat()
    ant_df = hist[hist["data"].astype(str) < hojeiso] if hist is not None and len(hist) else pd.DataFrame(columns=COLS_HIST)
    ant_df = ant_df.sort_values("data").groupby("base").tail(1).set_index("base") if len(ant_df) else ant_df
    sb = S["bases"].set_index("base") if S else pd.DataFrame()
    bb = B["bases"].set_index("base") if B else pd.DataFrame()
    se = S["emb_base"] if S else pd.DataFrame(columns=["base", "emb"])
    tk = S["tiktok"] if S else pd.DataFrame(columns=["base"])
    of = B["ofensores"] if B else pd.DataFrame(columns=["base"])
    grupos, links = grupos or {}, links or {}
    msgs = []
    for base in sorted(set(sb.index) | set(bb.index)):
        s = sb.loc[base].to_dict() if base in sb.index else None
        b = bb.loc[base].to_dict() if base in bb.index else None
        e = {r["emb"]: r for r in se[se["base"] == base].to_dict("records")}
        ant = ant_df.loc[base].to_dict() if base in getattr(ant_df, "index", []) else {}
        lista = [{"waybill": r["waybill"], "cidade": r["cidade"], "motorista": r["motorista"], "em_rota": bool(r["em_rota"]),
                  "diff": None if pd.isna(r["diff"]) else int(r["diff"]), "dias": int(r["dias"])}
                 for r in tk[tk["base"] == base].to_dict("records")]
        ofs = [(r["motorista"], int(r["qtd"]), float(r["valor"])) for r in of[of["base"] == base].head(5).to_dict("records")]
        txt = mensagem.montar(base, data_txt, ontem_txt, s, e, b, ant, lista, ofs, links.get(base))
        g = grupos.get(base, {})
        msgs.append([base[:2], base, links.get(base, "⚠️ planilha não criada"), g.get("nome", ""), g.get("link", ""), txt,
                     (s or {}).get("sla_a_ontem"), (s or {}).get("sla_b_ontem"), (b or {}).get("total"), (b or {}).get("valor_perdas")])
    return msgs


def gerar(S, B, destino, hist_antigo=None, grupos=None, links=None, tirar_foto=True):
    """S/B = resultados de motor.calcular_sla / calcular_backlog (todos os estados juntos).
    tirar_foto=False (rodadas só de Backlog): o histórico fica como estava."""
    carimbo = f"SLA: {S['hoje']:%d/%m %H:%M} · Backlog: {B['hoje']:%d/%m %H:%M}"
    if tirar_foto or hist_antigo is None or hist_antigo.empty:
        hist = juntar_historico(hist_antigo, foto_do_dia(S, B))
    else:
        hist = hist_antigo
    msgs = montar_mensagens(S, B, hist, grupos, links)
    wb = Workbook()
    aba_resumo_estados(wb, S, B, carimbo)
    aba_whatsapp(wb, msgs, carimbo)
    aba_sla_bases(wb, S, carimbo)
    aba_sla_diario(wb, S, carimbo)
    aba_entradas(wb, S, carimbo)
    aba_embarcadores(wb, S, carimbo)
    aba_ranking(wb, S, carimbo)
    aba_cidades(wb, S, carimbo)
    aba_tiktok(wb, S, carimbo)
    aba_backlog(wb, B, carimbo)
    aba_perdas(wb, B, carimbo)
    aba_historico(wb, hist)
    Path(destino).parent.mkdir(parents=True, exist_ok=True)
    wb.save(destino)
    return destino, msgs, hist


def _juntar(resultados):
    """Junta os resultados de vários estados (cada estado é calculado com o seu BD)."""
    out = {}
    for k, v in resultados[0].items():
        out[k] = pd.concat([r[k] for r in resultados], ignore_index=True) if isinstance(v, pd.DataFrame) else v
    return out


def _ler(p):
    p = Path(p)
    return pd.read_parquet(p) if p.suffix == ".parquet" else pd.read_excel(p)


