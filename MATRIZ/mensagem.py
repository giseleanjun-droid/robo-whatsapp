# -*- coding: utf-8 -*-
"""
mensagem.py - mensagem única de WhatsApp por base: SLA (boletim operacional) + BACKLOG (relatório)
+ TIKTOK (prioridade), com a comparação com o dia anterior (melhorou / piorou).
"""
import math

MAX_AJ = 150   # teto de códigos TikTok listados (o resto fica na planilha)


def _br(n, casas=0):
    s = f"{n:,.{casas}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def pct(v):
    return "—" if v is None or (isinstance(v, float) and math.isnan(v)) else _br(v * 100, 1) + "%"


def rs(v):
    return "R$ " + _br(v or 0, 2)


def _vazio(v):
    return v is None or (isinstance(v, float) and math.isnan(v))


def tend_pct(atual, antes):
    """SLA: subir é bom."""
    if _vazio(atual) or _vazio(antes):
        return ""
    d = (atual - antes) * 100
    if abs(d) < 0.05:
        return " ⚪ igual"
    return f" 🟢 melhorou +{_br(d, 1)} pp" if d > 0 else f" 🔴 piorou {_br(d, 1)} pp"


def tend_qtd(atual, antes, txt_unid=""):
    """Backlog / atrasados / perdas: subir é ruim."""
    if _vazio(antes):
        return ""
    d = int(atual) - int(antes)
    if d == 0:
        return " ⚪ igual"
    return f" 🔴 +{_br(d)}{txt_unid}" if d > 0 else f" 🟢 {_br(d)}{txt_unid}"


def montar(base, data_txt, ontem_txt, s, se, b, ant, tiktok, ofensores, link):
    """
    s   = linha do SLA da base (dict) ou None
    se  = {emb: linha} SLA dos destaques da base
    b   = linha do Backlog da base (dict) ou None
    ant = foto do dia anterior (dict) ou {}
    tiktok = lista de dicts (pacotes TikTok em aberto: waybill, cidade, motorista, em_rota, diff, dias)
    ofensores = lista de (motorista, qtd, valor)
    """
    m = f"*📊 BOLETIM DIÁRIO — {base} · {data_txt}*\n\n"
    if link:   # link único da planilha (SLA + Backlog) no topo; a planilha é fechada: a base pede acesso
        m += f"🔗 *Sua planilha SLA + Backlog:*\n{link}\n_(se pedir acesso, clique em \"Solicitar acesso\")_\n\n"
    m += f"⬇️ *RESUMO DE ONTEM ({ontem_txt}) E PENDÊNCIAS DE HOJE*\n\n"

    if s:
        m += "📊 *PERFORMANCE SLA ANJUN*\n"
        m += f"✅ Ontem ({ontem_txt}): *{pct(s['sla_a_ontem'])}*{tend_pct(s['sla_a_ontem'], ant.get('sla_a_ontem'))}\n"
        m += f"🏆 Mês (MTD): *{pct(s['sla_a_mes'])}*\n\n"
        m += "📈 *PERFORMANCE SLA BASE*\n"
        m += f"✅ Ontem ({ontem_txt}): *{pct(s['sla_b_ontem'])}*{tend_pct(s['sla_b_ontem'], ant.get('sla_b_ontem'))}\n"
        m += f"🏆 Mês (MTD): *{pct(s['sla_b_mes'])}*\n\n"
        m += "📦 *PRAZO DOS PACOTES (SLA)*\n"
        m += f"🔴 Atrasados: *{_br(s['atrasado'])}* (Foco Total!){tend_qtd(s['atrasado'], ant.get('atrasado'))}\n"
        m += f"🟡 Vence Hoje: *{_br(s['hoje'])}*\n"
        m += f"🟢 Amanhã: {_br(s['amanha'])} · 🔵 Semana: {_br(s['semana'])}\n"
        m += f"🏢 Na Base: {_br(s['na_base'])} · 🏍️ Em Rota: {_br(s['em_rota'])}\n\n"

    if b:
        m += "📦 *BACKLOG*\n"
        m += f"Total pendente: *{_br(b['total'])}*{tend_qtd(b['total'], ant.get('bl_total'))}\n"
        m += f"💰 Valor em risco: {rs(b['valor'])}\n"
        m += f"🔥 Vencidos: *{_br(b['vencidos'])}* · ⚠️ Vence hoje: *{_br(b['vence_hoje'])}*\n"
        m += f"🚨 Mais de 7 dias (extravio): *{_br(b['extravio'])}*{tend_qtd(b['extravio'], ant.get('bl_extravio'))}\n"
        m += f"🏍️ Em Rota: {_br(b['rota'])} · 📦 Estoque: {_br(b['estoque'])} · 🌍 Outra região: {_br(b['outra_regiao'] + b['outros'])}\n\n"
        m += "⏳ *TEMPO PARADO*\n"
        m += (f"🔴 > 7d: {_br(b['age_mais7'])} · 🟠 7d: {_br(b['age7'])} · 🟠 6d: {_br(b['age6'])}\n"
              f"🟡 5d: {_br(b['age5'])} · 4d: {_br(b['age4'])} · 3d: {_br(b['age3'])} · 2d: {_br(b['age2'])} · 🟢 1d: {_br(b['age1'])}\n\n")
        m += f"💸 *PERDA CONFIRMADA:* {rs(b['valor_perdas'])} ({_br(b['perdas'])} pacotes)"
        m += tend_qtd(b["perdas"], ant.get("bl_perdas"), " pacote(s)") + "\n\n"

    m += _bloco_tiktok(se.get("TIKTOK"), tiktok, ontem_txt)

    if ofensores:
        m += "🚨 *TOP OFENSORES (EM ROTA)*\n"
        for nome, q, v in ofensores[:5]:
            m += f"- {nome}: {_br(q)} ({rs(v)})\n"
        m += "\n"

    m += "⚠️ *AÇÃO:* Priorizar Atrasados, TikTok e +7 dias!"
    return m


def _bloco_tiktok(k, lista, ontem_txt):
    m = "🎵 *TIKTOK — PRIORIDADE*\n"
    if k:
        m += f"📊 SLA Anjun TikTok: ontem *{pct(k['sla_a_ontem'])}* · mês *{pct(k['sla_a_mes'])}*\n"
        m += f"📈 SLA Base TikTok: ontem *{pct(k['sla_b_ontem'])}* · mês *{pct(k['sla_b_mes'])}*\n"
    if not lista:
        return m + "✅ Nenhum pacote TikTok pendente na base ou em rota.\n\n"
    atr = sum(1 for x in lista if x["diff"] is not None and x["diff"] < 0)
    hj = sum(1 for x in lista if x["diff"] == 0)
    m += f"📦 *{len(lista)}* pendente(s) · 🔴 {atr} atrasado(s) · 🟡 {hj} vence(m) hoje\n"
    marca = lambda x: "⚠️" if x["dias"] > 7 else ("🔴" if (x["diff"] is not None and x["diff"] < 0) else ("🟡" if x["diff"] == 0 else "▫️"))
    usadas = 0
    rota = [x for x in lista if x["em_rota"]]
    base = [x for x in lista if not x["em_rota"]]
    if rota:
        m += f"\n🏍️ *EM ROTA ({len(rota)})*\n"
        por_mot = {}
        for x in rota:
            por_mot.setdefault(x["motorista"], []).append(x)
        for n, (mot, xs) in enumerate(por_mot.items()):
            if usadas >= MAX_AJ:
                break
            if n:
                m += "\n"   # linha em branco entre motoristas (fica fácil copiar o bloco de cada um)
            m += f"👤 *{mot}*\n"
            for x in xs:
                if usadas < MAX_AJ:
                    m += f"{marca(x)} {x['waybill']} · {x['cidade']}\n"
                usadas += 1
    if base and usadas < MAX_AJ:
        m += f"\n🏢 *NA BASE ({len(base)})*\n"
        for x in base:
            if usadas < MAX_AJ:
                mot = f" · 👤 {x['motorista']}" if x["motorista"] and x["motorista"] != "N/D" else ""
                m += f"{marca(x)} {x['waybill']} · {x['cidade']}{mot}\n"
            usadas += 1
    if len(lista) > MAX_AJ:
        m += f"… e mais {len(lista) - MAX_AJ} na sua planilha (aba 📋 SLA PACOTES, TikTok em vermelho)\n"
    return m + "🔴 atrasado · 🟡 vence hoje · ⚠️ +7 dias parado\n\n"
