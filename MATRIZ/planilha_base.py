# -*- coding: utf-8 -*-
"""
planilha_base.py - a planilha de UMA base (SLA + BACKLOG num arquivo só), no visual das antigas:

  📊 SLA RESULTADOS     painel de gestão (cards, gestão do dia, acumulado, dia a dia, motoristas, cidades)
  📋 SLA PACOTES        lista única: coluna 1 = 🔴 ATRASADO / 🟡 HOJE / 🟢 AMANHÃ / 🔵 SEMANA,
                        PRAZO BASE em destaque, linhas do TIKTOK destacadas
  📦 BACKLOG RESULTADOS painel de risco (total, perda, categorias, tempo parado, situação, motoristas)
  🚨 BACKLOG PACOTES    lista única: coluna 1 = 🚨 VENCIDO / 💸 PERDA / 🏍️ ROTA / 📦 ESTOQUE / 🌍 OUTRA REGIÃO
"""
import datetime as dt

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font
from openpyxl.worksheet.datavalidation import DataValidation

RESPOSTAS_CLIENTE = ["✅ SIM - recebeu", "❌ NÃO - não recebeu", "🚚 Entregue hoje", "📵 Não atende", "🔴 Número errado"]

import motor
from matriz import (_escrever, _semaforo, _texto_regras, FILL, ESCURO, BRANCO, VERDE, AMARELO, VERMELHO,
                    CINZA, F_PCT, F_INT, F_RS, F_DATA_H, F_DATA)
from openpyxl.utils import get_column_letter

COR_TK = "CC0000"   # TIKTOK: só a célula do embarcador, em vermelho


def _centralizar(ws, ini, fim, ncols):
    """Tudo centralizado (pedido da Gisele); o texto não quebra linha para a lista ficar compacta."""
    al = Alignment(horizontal="center", vertical="center")
    for row in ws.iter_rows(min_row=ini, max_row=fim, max_col=ncols):
        for c in row:
            c.alignment = al


def _tiktok_vermelho(ws, cols, n):
    """Só a célula 'TIKTOK' da coluna Embarcador fica vermelha (antes a linha inteira ficava preta)."""
    L = get_column_letter(cols.index("Embarcador") + 1)
    ws.conditional_formatting.add(f"{L}3:{L}{n}", FormulaRule(
        formula=[f'{L}3="{motor.PRIORIDADE}"'], fill=FILL(COR_TK), font=Font(color=BRANCO, bold=True)))


def _cor_sla(v):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return CINZA
    return VERDE if v >= 0.98 else (AMARELO if v >= 0.8 else VERMELHO)


def _pct(v):
    return "—" if v is None or (isinstance(v, float) and np.isnan(v)) else f"{v * 100:.1f}%".replace(".", ",")


def _bloco(ws, lin, titulo, cab, linhas, cor=ESCURO, formatos=None):
    """Título + cabeçalho + linhas a partir da coluna A. Devolve a próxima linha livre."""
    ws.cell(lin, 1, titulo).font = Font(bold=True, size=12)
    for j, c in enumerate(cab, 1):
        x = ws.cell(lin + 1, j, c)
        x.fill, x.font = FILL(cor), Font(bold=True, color=BRANCO)
        x.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for i, r in enumerate(linhas, lin + 2):
        for j, v in enumerate(r, 1):
            if v is None or (isinstance(v, float) and np.isnan(v)):
                continue
            c = ws.cell(i, j, v.item() if hasattr(v, "item") else v)
            c.alignment = Alignment(horizontal="center", vertical="center")
            if formatos and j in formatos:
                c.number_format = formatos[j]
    return lin + len(linhas) + 3


def _card(ws, lin, col, texto, cor):
    ws.merge_cells(start_row=lin, start_column=col, end_row=lin + 1, end_column=col + 1)
    c = ws.cell(lin, col, texto)
    c.fill, c.font = FILL(cor), Font(bold=True, size=13)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


# ------------------------------------------------------------------ SLA
def aba_sla_resultados(wb, base, S, carimbo):
    ws = wb.active
    ws.title = "📊 SLA RESULTADOS"
    ws.sheet_view.showGridLines = False
    ws.cell(1, 1, carimbo).font = Font(size=8, color="666666")
    bs = S["bases"]
    if base not in set(bs["base"]):
        ws.cell(3, 1, "Sem pacotes de SLA desta base no mês.")
        return
    k = bs[bs["base"] == base].iloc[0]
    eb = S["emb_base"][S["emb_base"]["base"] == base].set_index("emb")
    e = lambda emb, c: eb.at[emb, c] if emb in eb.index else 0

    _card(ws, 2, 1, f"SLA MÊS ANJUN\n{_pct(k['sla_a_mes'])}", _cor_sla(k["sla_a_mes"]))
    _card(ws, 2, 3, f"SLA MÊS BASE\n{_pct(k['sla_b_mes'])}", _cor_sla(k["sla_b_mes"]))
    if "sla_m_mes" in k:
        _card(ws, 2, 9, f"📏 SLA MALHA (km)\n{_pct(k['sla_m_mes'])}", _cor_sla(k["sla_m_mes"]))
    for i, emb in enumerate(motor.DESTAQUES):
        a = eb.at[emb, "sla_a_mes"] if emb in eb.index else np.nan
        b = eb.at[emb, "sla_b_mes"] if emb in eb.index else np.nan
        _card(ws, 2, 5 + i * 2, f"{emb} · SLA MÊS\nAnjun {_pct(a)} · Base {_pct(b)}", _cor_sla(a))
    lin = 5

    dia = f"{S['hoje']:%d/%m}"
    cab = [dia, "PACOTES DO DIA (ANJUN)", "✅ SUCESSO", "❌ FALHA", "📦 BACKLOG", "🔴 ATRASADO", "🟡 HOJE",
           "🟢 AMANHÃ", "🔵 SEMANA", "🏢 NA BASE", "🏍️ EM ROTA", "🚨 +7 DIAS"]
    linhas = [["GERAL", k["dem_hoje"], k["suc_hoje"], k["dem_hoje"] - k["suc_hoje"], k["backlog"], k["atrasado"], k["hoje"],
               k["amanha"], k["semana"], k["na_base"], k["em_rota"], k["extravio"]]]
    for emb in motor.DESTAQUES:
        linhas.append([emb, e(emb, "dem_hoje"), e(emb, "suc_hoje"), e(emb, "dem_hoje") - e(emb, "suc_hoje"), e(emb, "backlog"),
                       e(emb, "atrasado"), e(emb, "hoje"), e(emb, "amanha"), e(emb, "semana"), e(emb, "na_base"),
                       e(emb, "em_rota"), e(emb, "extravio")])
    lin = _bloco(ws, lin, "Gestão do dia da Base 🚀", cab, linhas)

    cab = ["TOTAL", "PACOTES DO MÊS (ANJUN)", "✅ SUCESSO", "❌ FALHA", "SLA ANJUN", "SLA BASE", "SLA ANJUN ONTEM",
           "SLA BASE ONTEM", "CHEGOU VENCIDO", "📏 SLA MALHA (km)", "📏 MALHA ONTEM"]
    linhas = [["GERAL", k["dem_mes"], k["suc_mes"], k["dem_mes"] - k["suc_mes"], k["sla_a_mes"], k["sla_b_mes"],
               k["sla_a_ontem"], k["sla_b_ontem"], k["cv_mes"], k.get("sla_m_mes"), k.get("sla_m_ontem")]]
    for emb in motor.DESTAQUES:
        if emb in eb.index:
            r = eb.loc[emb]
            linhas.append([emb, r["dem_mes"], r["suc_mes"], r["dem_mes"] - r["suc_mes"], r["sla_a_mes"], r["sla_b_mes"],
                           r["sla_a_ontem"], r["sla_b_ontem"], r["cv_mes"], r.get("sla_m_mes"), r.get("sla_m_ontem")])
    ini = lin + 2
    lin = _bloco(ws, lin, "Acumulado do mês 🚀", cab, linhas, formatos={5: F_PCT, 6: F_PCT, 7: F_PCT, 8: F_PCT, 10: F_PCT, 11: F_PCT})
    _semaforo(ws, f"E{ini}:H{ini + len(linhas) - 1}")
    _semaforo(ws, f"J{ini}:K{ini + len(linhas) - 1}")

    # dia a dia
    d = S["diario"]
    d = d[d["base"] == base].set_index("dia")
    dias = list(range(1, S["hoje"].day + 1))
    g = lambda c, x: (d.at[x, c] if (x in d.index and c in d.columns) else 0)
    def taxa(a, b, x):
        return (g(a, x) / g(b, x)) if g(b, x) else None
    linhas = [["SLA Anjun %"] + [taxa("suc", "dem", x) for x in dias],
              ["SLA Base %"] + [taxa("sucb", "demb", x) for x in dias],
              ["📏 SLA Malha (km) %"] + [taxa("sucm", "demm", x) for x in dias]]
    for emb in motor.DESTAQUES:
        linhas.append([f"{emb} · SLA Anjun %"] + [taxa(f"suc_{emb}", f"dem_{emb}", x) for x in dias])
        linhas.append([f"{emb} · SLA Base %"] + [taxa(f"sucb_{emb}", f"demb_{emb}", x) for x in dias])
    n_pct = len(linhas)
    for rot, c in [("Pacotes do dia SLA Anjun", "dem"), ("Não entregues SLA Anjun", None), ("Entregue (Baixas)", "entregue"),
                   ("Recebido", "recebido"), (f"Recebido após {motor.HORA_CORTE:02d}h", "recebido_apos9"),
                   ("Chegou atrasado do dia", "chegou_atrasado"), ("Pacotes com problema", "problema")]:
        if c is None:
            linhas.append([rot] + [(g("dem", x) - g("suc", x)) or None for x in dias])
        else:
            linhas.append([rot] + [g(c, x) or None for x in dias])
    ini = lin + 2
    cab = ["DIA A DIA (Meta 98%)"] + [f"{x:02d}/{S['hoje'].month:02d}" for x in dias]
    lin = _bloco(ws, lin, "Dia a dia performance SLA 🚀", cab, linhas)
    for i in range(ini, ini + n_pct):
        for j in range(2, len(dias) + 2):
            ws.cell(i, j).number_format = "0%"
        _semaforo(ws, f"B{i}:{get_column_letter(len(dias) + 1)}{i}")

    # motoristas e cidades
    m = S["motoristas"]
    m = m[(m["base"] == base) & ((m["entregues"] > 0) | (m["backlog"] > 0))].sort_values("fora_b", ascending=False).head(20)
    cab = ["MOTORISTA", "VALOR ABERTO", "ENTREGUES", "SLA ANJUN %", "SLA BASE %", "BACKLOG", "ATRASADOS",
           "TIKTOK ENTREGUES", "TIKTOK BACKLOG"]
    linhas = m[["motorista", "valor_aberto", "entregues", "sla_a", "sla_b", "backlog", "atrasado",
                "TIKTOK_entregues", "TIKTOK_backlog"]].values.tolist()
    ini = lin + 2
    lin = _bloco(ws, lin, "🚚 MOTORISTAS OFENSORES EM SLA BASE (volume)", cab, linhas, formatos={2: F_RS, 4: F_PCT, 5: F_PCT})
    if linhas:
        _semaforo(ws, f"D{ini}:E{ini + len(linhas) - 1}")
    c = S["cidades"]
    c = c[(c["base"] == base) & ((c["entregues"] > 0) | (c["backlog"] > 0))].copy()
    c["fails"] = c["fora_a"] + c["fora_b"]
    c = c.sort_values("fails", ascending=False).head(15)
    cab = ["CIDADE", "ENTREGUES", "SLA ANJUN %", "SLA BASE %", "BACKLOG", "ATRASADOS", "CHEGOU VENCIDO", "FAILS TOTAIS"]
    linhas = c[["cidade", "entregues", "sla_a", "sla_b", "backlog", "atrasado", "chegou_vencido", "fails"]].values.tolist()
    ini = lin + 2
    _bloco(ws, lin, "🏙️ CIDADES OFENSORAS SLA BASE e ANJUN", cab, linhas, formatos={3: F_PCT, 4: F_PCT})
    if linhas:
        _semaforo(ws, f"C{ini}:D{ini + len(linhas) - 1}")
    ws.column_dimensions["A"].width = 30
    for j in range(2, 40):
        ws.column_dimensions[get_column_letter(j)].width = 12


def aba_sla_pacotes(wb, base, S, carimbo):
    ws = wb.create_sheet("📋 SLA PACOTES")
    l = S["lista"]
    l = l[l["BASE"] == base].drop(columns=[c for c in l.columns if c.startswith("_")] + ["UF", "BASE"])
    cont = l["PRAZO ANJUN"].value_counts()
    ws.cell(1, 1, f"{carimbo}  ·  🔴 {cont.get(motor.PZ_ATRASADO, 0)} atrasado · 🟡 {cont.get(motor.PZ_HOJE, 0)} hoje · "
                  f"🟢 {cont.get(motor.PZ_AMANHA, 0)} amanhã · 🔵 {cont.get(motor.PZ_SEMANA, 0)} semana  ·  "
                  f"TIKTOK em vermelho").font = Font(size=9, bold=True)
    ws.sheet_view.showGridLines = False
    if l.empty:
        ws.cell(3, 1, "Zerado! ✅")
        return
    _escrever(ws, l, formatos={"Valor": F_RS, "Recebido na base": F_DATA_H, "Data Custódia": F_DATA_H,
                               "Data Anjun": F_DATA_H, "Data Base": F_DATA, "Reg. Ocorrência": F_DATA_H},
              larguras={"Waybill": 20, "Motivo": 32, "Motorista": 22, "Cidade": 20, "VALIDAÇÃO": 22})
    n = len(l) + 2
    cols = list(l.columns)
    _centralizar(ws, 3, n, len(cols))
    # cor só nos avisos: prazo (Anjun e Base) e a célula TIKTOK
    _texto_regras(ws, f"A3:B{n}", [("ATRASADO", "CC0000", BRANCO), ("HOJE", "FF9900", None),
                                   ("AMANH", "D9EAD3", None), ("SEMANA", "CFE2F3", None)])
    _tiktok_vermelho(ws, cols, n)


# ------------------------------------------------------------------ BACKLOG
def aba_bl_resultados(wb, base, B, carimbo):
    ws = wb.create_sheet("📦 BACKLOG RESULTADOS")
    ws.sheet_view.showGridLines = False
    ws.cell(1, 1, carimbo).font = Font(size=8, color="666666")
    bb = B["bases"]
    if base not in set(bb["base"]):
        ws.cell(3, 1, "Backlog zerado! ✅")
        return
    k = bb[bb["base"] == base].iloc[0]
    ws.merge_cells("A2:D3")
    c = ws.cell(2, 1, f"BACKLOG TOTAL: {int(k['total'])}")
    c.fill, c.font = FILL(ESCURO if k["total"] else "34A853"), Font(bold=True, size=22, color=BRANCO)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells("A5:D5")
    c = ws.cell(5, 1, f"💸 PERDA CONFIRMADA = R$ {k['valor_perdas']:,.2f}  ({int(k['perdas'])} pacotes)"
                .replace(",", "X").replace(".", ",").replace("X", "."))
    c.font, c.alignment = Font(bold=True, color="CC0000", size=12), Alignment(horizontal="center")
    ws.merge_cells("A6:B6")
    c = ws.cell(6, 1, f"⚠️ deste mês (vai descontar): R$ {k.get('valor_perdas_mes', 0):,.2f} ({int(k.get('perdas_mes', 0))} pct)"
                .replace(",", "X").replace(".", ",").replace("X", "."))
    c.font, c.alignment = Font(bold=True, color="CC0000", size=10), Alignment(horizontal="center")
    ws.merge_cells("C6:D6")
    c = ws.cell(6, 3, f"✅ já reembolsado (mês anterior - não cobre o cliente): R$ {k.get('valor_perdas_ant', 0):,.2f} ({int(k.get('perdas_ant', 0))} pct)"
                .replace(",", "X").replace(".", ",").replace("X", "."))
    c.font, c.alignment = Font(bold=True, color="34A853", size=10), Alignment(horizontal="center")
    lin = 8
    lin = _bloco(ws, lin, "STATUS ATUAL", ["CATEGORIA", "QTD", "VALOR (Est.)"],
                 [["🏍️ EM ROTA", k["rota"], k["valor_rota"]], ["📦 ESTOQUE", k["estoque"], k["valor_estoque"]],
                  ["🌍 OUTRA REGIÃO", k["outra_regiao"], k["valor_outra"]], ["⚠️ OUTROS", k["outros"], k["valor_outros"]]],
                 formatos={3: F_RS})
    ini = lin + 2
    lin = _bloco(ws, lin, "TEMPO PARADO", ["TEMPO PARADO", "QTD", "STATUS"],
                 [["> 7 DIAS", k["age_mais7"], "EXTRAVIO?"], ["7 DIAS", k["age7"], "CRÍTICO"], ["6 DIAS", k["age6"], "CRÍTICO"],
                  ["5 DIAS", k["age5"], "ATENÇÃO"], ["4 DIAS", k["age4"], "ATENÇÃO"], ["3 DIAS", k["age3"], "ATENÇÃO"],
                  ["2 DIAS", k["age2"], "ATENÇÃO"], ["1 DIA", k["age1"], "NORMAL"]])
    for off, cor, fonte in [(0, "000000", BRANCO), (1, "FF9900", None), (2, "FF9900", None), (3, "F1C232", None),
                            (4, "F1C232", None), (5, "F1C232", None), (6, "F1C232", None), (7, "34A853", BRANCO)]:
        for j in range(1, 4):
            ws.cell(ini + off, j).fill = FILL(cor)
            if fonte:
                ws.cell(ini + off, j).font = Font(color=fonte, bold=True)
    ini = lin + 2
    lin = _bloco(ws, lin, "SITUAÇÃO", ["SITUAÇÃO", "QTD", "STATUS"],
                 [["🔥 VENCIDOS", k["vencidos"], "URGENTE"], ["⚠️ VENCE HOJE", k["vence_hoje"], "PRIORIDADE"],
                  ["📦 NO PRAZO", k["no_prazo"], "CONTROLADO"], ["🏴‍☠️ EXTRAVIO (+7d)", k["extravio"], "INVESTIGAR"]])
    for off, cor, fonte in [(0, "CC0000", BRANCO), (1, "FF9900", None), (2, "34A853", BRANCO), (3, "000000", BRANCO)]:
        for j in range(1, 4):
            ws.cell(ini + off, j).fill = FILL(cor)
            if fonte:
                ws.cell(ini + off, j).font = Font(color=fonte, bold=True)
    of = B["ofensores"]
    of = of[of["base"] == base].head(10)
    _bloco(ws, lin, "MOTORISTAS (ROTA)", ["MOTORISTA", "QTD", "VALOR"], of[["motorista", "qtd", "valor"]].values.tolist(),
           formatos={3: F_RS})
    for col, w in zip("ABCD", (30, 14, 16, 14)):
        ws.column_dimensions[col].width = w


def aba_bl_pacotes(wb, base, B, carimbo, anterior=None):
    ws = wb.create_sheet("🚨 BACKLOG PACOTES")
    l = B["lista"]
    l = l[l["BASE"] == base].drop(columns=["UF", "BASE"])
    cont = l["ALERTA"].value_counts()
    onde = l["ONDE ESTÁ"].value_counts()
    ws.cell(1, 1, f"{carimbo}  ·  " + " · ".join(f"{a} {cont.get(a, 0)}" for a in motor.ORDEM_AL if cont.get(a, 0)) +
            "  |  ONDE ESTÁ: " + " · ".join(f"{a} {onde.get(a, 0)}" for a in motor.ORDEM_AL if onde.get(a, 0)) +
            "  |  📞 contato com o cliente: pesquise o AJ no SRM, cole o telefone na coluna e o link do WhatsApp " +
            "já sai com a mensagem pronta  |  💸 Perda / Reembolso: se já foi reembolsado em mês anterior, NÃO cobre o cliente de novo"
            ).font = Font(size=9, bold=True)
    ws.sheet_view.showGridLines = False
    if l.empty:
        ws.cell(3, 1, "Zerado! ✅")
        return
    _escrever(ws, l, formatos={"Valor": F_RS, "Recebido na base": F_DATA_H, "Data Custódia": F_DATA_H,
                               "SLA Anjun": F_DATA_H, "SLA Base": F_DATA_H, "Reg. Ocorrência": F_DATA_H},
              larguras={"ALERTA": 22, "💸 Perda / Reembolso": 38, "ONDE ESTÁ": 18, "Waybill": 20, "Motivo": 34, "Motorista": 22, "Cidade": 20,
                       "VALIDAÇÃO": 22, "Cliente": 22, "Telefone (cole aqui)": 20,
                       "Mensagem pronta (copiar)": 40, "📲 Link WhatsApp": 16, "Resposta do cliente": 20})
    n = len(l) + 2
    cols = list(l.columns)
    _centralizar(ws, 3, n, len(cols))
    # cor só nos avisos: a coluna ALERTA e a célula TIKTOK
    _texto_regras(ws, f"A3:A{n}", [("VENCIDO", "CC0000", BRANCO), ("PERDA", "000000", BRANCO)])
    _texto_regras(ws, f"B3:B{n}", [("este mês", "CC0000", BRANCO), ("reembolsado", "34A853", BRANCO)])
    c_onde = get_column_letter(cols.index("ONDE ESTÁ") + 1)
    _texto_regras(ws, f"{c_onde}3:{c_onde}{n}", [("ROTA", "CFE2F3", None), ("ESTOQUE", "FFF2CC", None), ("OUTRA", "EAD1DC", None)])
    _tiktok_vermelho(ws, cols, n)
    _contato_whatsapp(ws, cols, n, anterior or {})
    _resposta_cliente(ws, cols, n, anterior or {})


def _contato_whatsapp(ws, cols, n, anterior):
    """Coluna 'Telefone (cole aqui)': o número NUNCA vem do robô (é dado protegido no SRM/Express - só aparece
    pessoa a pessoa, com verificação por código; não dá para automatizar, nem deve). A base pesquisa o AJ no
    SRM (acesso que ela já tem) e cola o telefone aqui. A fórmula ao lado monta sozinha o link do WhatsApp já
    com a mensagem pronta (igual ao padrão manual que a Gisele já usa: https://wa.me/<numero>?text=<mensagem>).

    O robô reescreve esta planilha a cada regravação (3x ao dia): sem isso, o telefone colado pela base se
    perderia a cada vez. `anterior` é {waybill: telefone} lido da planilha anterior (ver bases.py) para o
    telefone já digitado acompanhar o pacote mesmo com a lista remontada do zero.
    """
    if "Telefone (cole aqui)" not in cols or "Mensagem pronta (copiar)" not in cols or "📲 Link WhatsApp" not in cols:
        return
    c_tel = get_column_letter(cols.index("Telefone (cole aqui)") + 1)
    c_msg = get_column_letter(cols.index("Mensagem pronta (copiar)") + 1)
    j_wb, j_tel, j_link = cols.index("Waybill") + 1, cols.index("Telefone (cole aqui)") + 1, cols.index("📲 Link WhatsApp") + 1
    ws.cell(2, j_tel).comment = Comment(
        "Pesquise o AJ no SRM (Consulta de pedidos) e cole aqui só os números: "
        "DDI + DDD + número, sem espaço/traço. Exemplo: 5562984501234", "Anjun")
    for i in range(3, n + 1):
        wb_ = ws.cell(i, j_wb).value
        tel = (anterior.get(wb_) or {}).get("telefone")
        if tel:
            ws.cell(i, j_tel, tel)
        ws.cell(i, j_link, f'=IF({c_tel}{i}="","",HYPERLINK("https://wa.me/"&{c_tel}{i}&"?text="&ENCODEURL({c_msg}{i}),"📲 Abrir WhatsApp"))')


def _resposta_cliente(ws, cols, n, anterior=None):
    """Coluna 'Resposta do cliente': lista suspensa com o que a base apurou depois de mandar a mensagem.
    Preserva a resposta já marcada antes (igual ao telefone - ver _contato_whatsapp) para não perder o
    que já foi apurado a cada regravação automática da planilha.
    """
    if "Resposta do cliente" not in cols:
        return
    j_wb, j_resp = cols.index("Waybill") + 1, cols.index("Resposta do cliente") + 1
    c_resp = get_column_letter(j_resp)
    dv = DataValidation(type="list", formula1='"' + ",".join(RESPOSTAS_CLIENTE) + '"', allow_blank=True, showDropDown=False)
    ws.add_data_validation(dv)
    dv.add(f"{c_resp}3:{c_resp}{n}")
    anterior = anterior or {}
    for i in range(3, n + 1):
        wb_ = ws.cell(i, j_wb).value
        resp = (anterior.get(wb_) or {}).get("resposta")
        if resp:
            ws.cell(i, j_resp, resp)


def gerar(base, S, B, destino, anterior=None):
    carimbo = f"📅 SLA atualizado {S['hoje']:%d/%m %H:%M} · Backlog atualizado {B['hoje']:%d/%m %H:%M}"
    wb = Workbook()
    aba_sla_resultados(wb, base, S, carimbo)
    aba_sla_pacotes(wb, base, S, carimbo)
    aba_bl_resultados(wb, base, B, carimbo)
    aba_bl_pacotes(wb, base, B, carimbo, anterior)
    wb.save(destino)
    return destino
