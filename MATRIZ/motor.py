# -*- coding: utf-8 -*-
"""
motor.py - calcula SLA e BACKLOG de todas as bases a partir dos BD_<UF> (mesmas regras do
Código.gs do SLA e do BACKLOG, com as correções combinadas em 14/11/2026).

Tudo em Python (sem limite de linhas do Google Sheets). Só os resumos e os pacotes em aberto
vão para as planilhas.

Correções em relação aos scripts antigos:
  - Horário de Brasília em tudo (o SLA.gs estava com o fuso errado: +5h, 23:59 virava 04:59 do dia seguinte).
  - Extravio = MAIS DE 7 dias parado (SLA usava 10+, Backlog >7).
  - Dias parado = dias corridos desde a entrada no ponto (entrou ontem = 1), igual nos dois.
  - Perda = motivo/status/indenização com PERDA, DEBIT, DESCONT ou vencido há 20+ dias (regra do Backlog) nos dois.
  - Valor declarado em texto "1.234,56" lido certo.
"""
import re
import unicodedata
import datetime as dt
from zoneinfo import ZoneInfo

import numpy as np
import pandas as pd

FUSO = ZoneInfo("America/Sao_Paulo")
DIAS_EXTRAVIO = 7            # parado há MAIS de 7 dias
DIAS_VENCIDO_VIRA_PERDA = 20
PRAZO_BASE_PADRAO_DIAS = 2
HORA_CORTE = 9               # entrou às 09:00 ou depois -> SLA Base +1 dia (e conta como "recebido após o horário")
DESTAQUES = ["SHEIN", "TIKTOK"]
PRIORIDADE = "TIKTOK"
PADRAO_BASE = re.compile(r"^[A-Z]{2}-W-[A-Z]{1,2}\d+$")

FRASES_IGNORAR = [   # SLA: pacote conta nos números, mas não aparece na lista
    "O PACOTE CHEGOU AO CENTRO DE DISTRIBUIÇÃO", "O PACOTE ESTÁ A CAMINHO DO CENTRO DE TRÂNSITO",
    "O PACOTE FOI DEVOLVIDO AO CENTRO DE TRÂNSITO", "O PACOTE SAIU DE E FOI DEVOLVIDO AO CENTRO DE TRÂNSITO",
    "O PACOTE SAIU DO CENTRO DE TRANSBORDO", "O PACOTE SERÁ DEVOLVIDO AO VENDEDOR",
    "O PACOTE FOI COLETADO COM SUCESSO", "O PACOTE FOI ASSINADO [DEVOLUÇÃO ASSINADA]", "[SP-RR-001]",
]

REGRAS_EMB = [   # iguais ao Embarcadores.gs / evolucao_mensal.py: 1ª coluna que casar vence
    ("SHEIN", r"SHEIN"), ("TIKTOK", r"TIK\s*[-_]?\s*TOK"), ("KWAI", r"KWAI"), ("TEMU", r"TEMU"),
    ("SHOPEE", r"SHOPEE"), ("MERCADO LIVRE", r"MERCADO\s*LIVRE|MERCADOLIVRE"), ("AMAZON", r"AMAZON"),
    ("ALIEXPRESS", r"ALI\s*EXPRESS|ALIBABA|HANGYU[\s_-]*AE|[（(]AE(?![A-Z])"), ("C&A", r"^(?:C&A|C&_|CEA)"),
    ("CASAS BAHIA", r"CASAS\s*BAHIA"), ("MAGALU", r"MAGALU|MAGAZINE\s*LUIZA"),
]

# rótulos (coluna 1 das listas)
PZ_ATRASADO, PZ_HOJE, PZ_AMANHA, PZ_SEMANA = "🔴 ATRASADO", "🟡 HOJE", "🟢 AMANHÃ", "🔵 SEMANA"
AL_PERDA, AL_URGENTE, AL_ROTA, AL_ESTOQUE, AL_OUTRA, AL_OUTROS = (
    "💸 PERDA CONFIRMADA", "🚨 VENCIDO - URGENTE", "🏍️ EM ROTA", "📦 ESTOQUE", "🌍 OUTRA REGIÃO", "⚠️ OUTROS")


def agora():
    """Agora em Brasília, sem fuso (as datas do BD já vêm em horário de Brasília)."""
    return dt.datetime.now(FUSO).replace(tzinfo=None)


# ---------------------------------------------------------------- leitura de colunas
def _norm(c):
    c = unicodedata.normalize("NFD", str(c).lower())
    return "".join(ch for ch in c if ch.isascii() and ch.isalnum())


def _col(df, *nomes):
    m = {_norm(c): c for c in df.columns}
    for n in nomes:
        if n in m:
            return df[m[n]]
    for n in nomes:   # "contém" (ex.: "Tempo de entrada  no ponto" com espaço duplo)
        for k, c in m.items():
            if n in k:
                return df[c]
    return pd.Series([None] * len(df), index=df.index, dtype="object")


def _txt(s):
    return s.astype("string").fillna("").str.strip()


def _sem_acento_upper(s):
    s = _txt(s).map(lambda v: unicodedata.normalize("NFD", v))
    return s.str.replace(r"[̀-ͯ]", "", regex=True).str.upper().str.replace(r"\s+", " ", regex=True).str.strip()


def _data(s):
    if pd.api.types.is_datetime64_any_dtype(s):
        return pd.to_datetime(s)
    t = s.astype("string").str.strip()
    iso = pd.to_datetime(t, format="%Y-%m-%d %H:%M:%S", errors="coerce")
    falta = iso.isna() & t.notna() & (t != "")
    if falta.any():
        iso[falta] = pd.to_datetime(t[falta], dayfirst=True, errors="coerce")
    return iso


def _valor(s):
    if pd.api.types.is_datetime64_any_dtype(s):   # coluna formatada como data na planilha (42,62 vira 11/02/1900)
        return ((s - pd.Timestamp("1899-12-30")).dt.total_seconds() / 86400).round(2).fillna(0.0)
    s = s.map(lambda v: (v - dt.datetime(1899, 12, 30)).total_seconds() / 86400 if isinstance(v, dt.datetime) else v)
    if pd.api.types.is_numeric_dtype(s):
        return s.fillna(0).astype(float)
    t = s.astype("string").fillna("").str.replace("R$", "", regex=False).str.strip()
    br = t.str.contains(",", regex=False)   # "1.234,56" -> 1234.56
    t = t.where(~br, t.str.replace(".", "", regex=False).str.replace(",", ".", regex=False))
    return pd.to_numeric(t, errors="coerce").fillna(0.0)


def embarcador(df):
    c1 = _sem_acento_upper(_col(df, "nomedocliente"))
    c2 = _sem_acento_upper(_col(df, "nomedoclientemerchant"))
    if (c1 == "").all() and (c2 == "").all():
        return pd.Series("N/D", index=df.index, dtype="string")
    out = pd.Series(pd.NA, index=df.index, dtype="string")
    for col in (c1, c2):
        for nome, rx in REGRAS_EMB:
            out[out.isna() & col.str.contains(rx, regex=True)] = nome
    resto = out.isna()
    out[resto] = c1[resto].where(c1[resto] != "", c2[resto]).str[:40].replace("", "N/D")
    return out


def _dias_corridos(d, hoje):
    """Dias de calendário entre a data d e hoje (entrou ontem = 1)."""
    return (pd.Timestamp(hoje.date()) - d.dt.normalize()).dt.days


def _prazo_base(prazo_b, entrada, cidade):
    """SLA Base vazio: entrada + dias mais comuns da cidade (aprendidos do próprio BD), senão 2."""
    dias = ((prazo_b - entrada).dt.total_seconds() / 86400).round()
    ok = dias.notna() & (dias > 0)
    if ok.any():
        moda = pd.DataFrame({"c": cidade[ok], "d": dias[ok]}).groupby("c")["d"].agg(lambda x: x.mode().iloc[0])
    else:
        moda = pd.Series(dtype=float)
    estim = entrada + pd.to_timedelta(cidade.map(moda).fillna(PRAZO_BASE_PADRAO_DIAS), unit="D")
    prazo_b = prazo_b.fillna(estim)
    # Regra de todos os estados: pacote que entra na base DEPOIS das 08:59 ganha +1 dia no SLA Base.
    # (o SRM não aplica: o "Horário de entrega solicitado" é sempre entrada + N dias no mesmo horário)
    return prazo_b + pd.to_timedelta(entrou_apos_corte(entrada).astype(int), unit="D")


def entrou_apos_corte(entrada):
    """Entrada no ponto a partir das 09:00 (horário de Brasília)."""
    h = entrada.dt.hour * 60 + entrada.dt.minute
    return (h >= HORA_CORTE * 60).fillna(False)


def _classe_prazo(d, hoje):
    """ATRASADO / HOJE / AMANHÃ / SEMANA pela DATA (Brasília) do prazo."""
    diff = (d.dt.normalize() - pd.Timestamp(hoje.date())).dt.days
    out = pd.Series(pd.NA, index=d.index, dtype="string")
    out[diff < 0] = PZ_ATRASADO
    out[diff == 0] = PZ_HOJE
    out[diff == 1] = PZ_AMANHA
    out[(diff >= 2) & (diff <= 7)] = PZ_SEMANA
    return out, diff


def _perda(status_n, motivo_n, indeniz_n, dias_vencido):
    return (motivo_n.str.contains("PERDA") | status_n.str.contains("PERDA") |
            indeniz_n.str.contains("PERDA|DEBIT|DESCONT", regex=True) |
            (dias_vencido >= DIAS_VENCIDO_VIRA_PERDA))


def _categoria(status_n, motivo_n):
    """ROTA / OUTRA REGIÃO / ESTOQUE / OUTROS (mesma ordem de teste dos scripts)."""
    cat = pd.Series("OUTROS", index=status_n.index, dtype="string")
    estoque = pd.Series(False, index=status_n.index)
    for k in ["RECEBIDO", "CHEGOU", "ARMAZENADO", "AGUARDANDO", "FALHA", "INSUCESSO", "PROBLEMA"]:
        estoque |= status_n.str.contains(k) | motivo_n.str.contains(k)
    cat[estoque] = "ESTOQUE"
    cat[status_n.str.contains("PERTENCE") | motivo_n.str.contains("PERTENCE") | motivo_n.str.contains("WRONG")] = "OUTRA REGIÃO"
    cat[status_n.str.contains("ROTA") | status_n.str.contains("SAIU") | status_n.str.contains("MOTORISTA")] = "ROTA"
    return cat


def _base_valida(base):
    return base.str.match(PADRAO_BASE)


def _comum(df, hoje):
    """Colunas padronizadas usadas pelo SLA e pelo Backlog."""
    t = pd.DataFrame(index=df.index)
    t["base"] = _txt(_col(df, "pontoresponsavel")).str.upper()
    t["uf"] = t["base"].str[:2]
    t["waybill"] = _txt(_col(df, "numerodowaybill", "waybill"))
    t["status"] = _txt(_col(df, "statusdopacote")).str.upper()
    t["status_n"] = _sem_acento_upper(_col(df, "statusdopacote"))
    t["motivo"] = _txt(_col(df, "motivodaocorrencia"))
    t["motivo_n"] = _sem_acento_upper(_col(df, "motivodaocorrencia"))
    t["indeniz_n"] = _sem_acento_upper(_col(df, "statusdaindenizacao"))
    t["motorista"] = _txt(_col(df, "entregador")).str.upper().replace("", "N/D")
    t["cidade"] = _txt(_col(df, "cidadedodestinatario")).str.upper().replace("", "N/D")
    t["valor"] = _valor(_col(df, "valordeclarado"))
    t["prazo"] = _data(_col(df, "horarioemquedeveserentregue"))
    t["entrada"] = _data(_col(df, "tempodeentradanoponto", "tempodeentrada"))
    t["hub"] = _txt(_col(df, "hubdetransito")).str.upper().replace("", "N/D")
    t["apos_corte"] = entrou_apos_corte(t["entrada"])
    t["custodia"] = _data(_col(df, "tempodeentreganoponto"))
    t["liberacao"] = _data(_col(df, "tempodeliberacaodoveiculo"))
    t["ocorrencia"] = _data(_col(df, "tempoderegistrodeocorrencia"))
    t["assinatura"] = _data(_col(df, "tempodeassinatura"))
    t["ult_rastreio"] = _data(_col(df, "ultimodataderastreio"))
    t["registrante"] = _txt(_col(df, "registrantedaocorrencia"))
    t["tent"] = _col(df, "totaldetentativas")
    t["rest"] = _col(df, "tentativasrestantes")
    t["rota_txt"] = _txt(_col(df, "conteudodaultimarota")).str.upper()
    t["emb"] = embarcador(df)
    t["prazo_b"] = _prazo_base(_data(_col(df, "horariodeentregasolicitado")), t["entrada"], t["cidade"])
    t["dias"] = _dias_corridos(t["entrada"].fillna(t["custodia"]).fillna(t["ult_rastreio"]), hoje).fillna(0).clip(lower=0).astype(int)
    t["pz_classe"], t["pz_diff"] = _classe_prazo(t["prazo"], hoje)
    t["pzb_classe"], _ = _classe_prazo(t["prazo_b"], hoje)
    venc = (-t["pz_diff"]).where(t["pz_diff"] < 0, 0).fillna(0)
    t["perda"] = _perda(t["status_n"], t["motivo_n"], t["indeniz_n"], venc)
    t["categoria"] = _categoria(t["status_n"], t["motivo_n"])
    t["em_rota"] = t["categoria"] == "ROTA"
    return t[_base_valida(t["base"])].copy()


def _validacao(t):
    """Texto da coluna VALIDAÇÃO (igual aos scripts, com extravio > 7 dias)."""
    v = pd.Series("", index=t.index, dtype="string")
    falta = t["pz_diff"].fillna(999).astype(int)
    v[:] = "⏳ Falta " + falta.astype(str) + " dias"
    v[t["pz_diff"].isna()] = "⏳ Sem prazo"
    v[falta == 0] = "⚠️ VENCE HOJE"
    v[falta < 0] = "🔥 ATRASADO"
    v[t["dias"] > DIAS_EXTRAVIO] = f"⚠️ EXTRAVIADO (> {DIAS_EXTRAVIO} dias)"
    return v


# ================================================================ SLA
def calcular_sla(df, hoje=None):
    """
    Devolve dict:
      bases:    DataFrame uma linha por base (KPIs do mês / ontem / backlog do dia)
      emb_base: DataFrame base x embarcador (destaques)
      emb_uf:   DataFrame UF x embarcador (todos)
      diario:   DataFrame base x dia (SLA Anjun/Base, entregues, recebidos, chegou atrasado, problemas)
      cidades, motoristas: ofensores por base
      lista:    pacotes em aberto (Atrasado / Hoje / Amanhã / Semana) com colunas da planilha
    """
    hoje = hoje or agora()
    ontem = (hoje - dt.timedelta(days=1)).date()
    ano, mes, dia_hoje = hoje.year, hoje.month, hoje.day
    t = _comum(df, hoje)

    entregue = t["status_n"].str.contains("ENTREGUE|ASSINADO", regex=True)
    devolvido = t["status_n"].str.contains("DEVOLVIDO|CANCELADO|RETORNA", regex=True)
    finalizado = entregue | devolvido
    entrega = t["assinatura"].where(t["assinatura"].notna() | ~entregue, t["ult_rastreio"])

    no_mes = lambda d: (d.dt.year == ano) & (d.dt.month == mes)
    ate_hoje = lambda d: no_mes(d) & (d.dt.day <= dia_hoje)
    eh_ontem = lambda d: d.dt.date == ontem

    a = t["prazo"].notna() & ~devolvido
    b = t["prazo_b"].notna() & ~devolvido
    ok_a = a & entregue & (entrega <= t["prazo"])
    ok_b = b & entregue & (entrega <= t["prazo_b"])
    chegou_venc = a & (t["entrada"] > t["prazo"])

    aberto = ~finalizado & t["entrada"].notna()
    aberto_ok = aberto & ~t["perda"]          # perda sai do backlog do SLA (igual antes)
    t["validacao"] = _validacao(t)

    f = pd.DataFrame({
        "base": t["base"], "uf": t["uf"], "emb": t["emb"],
        "dem_mes": a & ate_hoje(t["prazo"]), "suc_mes": ok_a & ate_hoje(t["prazo"]),
        "demb_mes": b & ate_hoje(t["prazo_b"]), "sucb_mes": ok_b & ate_hoje(t["prazo_b"]),
        "dem_ontem": a & eh_ontem(t["prazo"]), "suc_ontem": ok_a & eh_ontem(t["prazo"]),
        "demb_ontem": b & eh_ontem(t["prazo_b"]), "sucb_ontem": ok_b & eh_ontem(t["prazo_b"]),
        "dem_hoje": a & (t["prazo"].dt.date == hoje.date()), "suc_hoje": ok_a & (t["prazo"].dt.date == hoje.date()),
        "cv_mes": chegou_venc & ate_hoje(t["prazo"]), "cv_falha": chegou_venc & ate_hoje(t["prazo"]) & ~ok_a,
        "recebidos_mes": no_mes(t["entrada"]), "entregues_mes": entregue & no_mes(entrega),
        "backlog": aberto_ok, "valor_aberto": t["valor"].where(aberto_ok, 0.0),
        "atrasado": aberto_ok & (t["pz_classe"] == PZ_ATRASADO),
        "hoje": aberto_ok & (t["pz_classe"] == PZ_HOJE),
        "amanha": aberto_ok & (t["pz_classe"] == PZ_AMANHA),
        "semana": aberto_ok & (t["pz_classe"] == PZ_SEMANA),
        "na_base": aberto_ok & ~t["em_rota"], "em_rota": aberto_ok & t["em_rota"],
        "extravio": aberto_ok & (t["dias"] > DIAS_EXTRAVIO),
        "perdas": aberto & t["perda"], "valor_perdas": t["valor"].where(aberto & t["perda"], 0.0),
    })
    soma = [c for c in f.columns if c not in ("base", "uf", "emb")]

    bases = f.groupby(["uf", "base"])[soma].sum().reset_index()
    emb_base = f[f["emb"].isin(DESTAQUES)].groupby(["uf", "base", "emb"])[soma].sum().reset_index()
    emb_uf = f.groupby(["uf", "emb"])[soma].sum().reset_index()
    for d in (bases, emb_base, emb_uf):
        _pcts(d)

    # ---- dia a dia (pelo dia do prazo / entrada / entrega no mês)
    reg = []
    def por_dia(nome, mask, data):
        m = mask & no_mes(data)
        if m.any():
            reg.append(pd.DataFrame({"base": t.loc[m, "base"], "dia": data[m].dt.day, "k": nome}))
    por_dia("dem", a, t["prazo"]); por_dia("suc", ok_a, t["prazo"])
    por_dia("demb", b, t["prazo_b"]); por_dia("sucb", ok_b, t["prazo_b"])
    por_dia("entregue", entregue, entrega); por_dia("recebido", pd.Series(True, index=t.index), t["entrada"])
    por_dia("chegou_atrasado", t["entrada"] > t["prazo"], t["entrada"])
    por_dia("recebido_apos9", t["apos_corte"], t["entrada"])
    por_dia("problema", pd.Series(True, index=t.index), t["ocorrencia"])
    for e in DESTAQUES:
        por_dia(f"dem_{e}", a & (t["emb"] == e), t["prazo"]); por_dia(f"suc_{e}", ok_a & (t["emb"] == e), t["prazo"])
        por_dia(f"demb_{e}", b & (t["emb"] == e), t["prazo_b"]); por_dia(f"sucb_{e}", ok_b & (t["emb"] == e), t["prazo_b"])
    diario = (pd.concat(reg).groupby(["base", "dia", "k"]).size().unstack("k", fill_value=0).reset_index()
              if reg else pd.DataFrame(columns=["base", "dia"]))
    diario = diario[diario["dia"] <= dia_hoje]

    # ---- ofensores: motoristas e cidades
    p = pd.DataFrame({
        "uf": t["uf"], "base": t["base"], "motorista": t["motorista"], "cidade": t["cidade"],
        "entregues": entregue & no_mes(entrega),
        "ok_a": entregue & no_mes(entrega) & (entrega <= t["prazo"]),
        "ok_b": entregue & no_mes(entrega) & (entrega <= t["prazo_b"]),
        "backlog": aberto_ok, "atrasado": aberto_ok & (t["pz_classe"] == PZ_ATRASADO),
        "valor_aberto": t["valor"].where(aberto_ok, 0.0),
        "chegou_vencido": no_mes(t["entrada"]) & (t["entrada"] > t["prazo"]),
    })
    for e in DESTAQUES:
        p[f"{e}_entregues"] = p["entregues"] & (t["emb"] == e)
        p[f"{e}_ok"] = p["ok_a"] & (t["emb"] == e)
        p[f"{e}_backlog"] = p["backlog"] & (t["emb"] == e)
    nums = [c for c in p.columns if c not in ("uf", "base", "motorista", "cidade")]
    motoristas = p.groupby(["uf", "base", "motorista"])[nums].sum().reset_index()
    cidades = p.groupby(["uf", "base", "cidade"])[nums].sum().reset_index()
    for d in (motoristas, cidades):
        d["fora_a"] = d["entregues"] - d["ok_a"]
        d["fora_b"] = d["entregues"] - d["ok_b"]
        d["sla_a"] = np.where(d["entregues"] > 0, d["ok_a"] / d["entregues"].replace(0, 1), np.nan)
        d["sla_b"] = np.where(d["entregues"] > 0, d["ok_b"] / d["entregues"].replace(0, 1), np.nan)

    # ---- lista única de pacotes (atrasado / hoje / amanhã / semana)
    esconder = pd.Series(False, index=t.index)
    for fr in FRASES_IGNORAR:
        esconder |= t["rota_txt"].str.contains(fr, regex=False)
    m = aberto_ok & t["pz_classe"].notna() & ~esconder
    lista = _lista_sla(t[m])
    # TikTok em aberto (todos os prazos, igual à aba 🎵 TIKTOK antiga): mais urgente primeiro
    tk = t[aberto_ok & (t["emb"] == PRIORIDADE) & ~esconder].copy()
    tk["_o"] = tk["pz_diff"].fillna(999)
    tk = tk.sort_values(["base", "_o", "dias"], ascending=[True, True, False])
    tiktok = pd.DataFrame({"uf": tk["uf"], "base": tk["base"], "waybill": tk["waybill"], "cidade": tk["cidade"],
                           "motorista": tk["motorista"], "em_rota": tk["em_rota"], "dias": tk["dias"],
                           "diff": tk["pz_diff"], "prazo": tk["prazo"], "prazo_b": tk["prazo_b"],
                           "classe": tk["pz_classe"].fillna("⚪ +7 DIAS"), "classe_b": tk["pzb_classe"].fillna("⚪ +7 DIAS"),
                           "status": tk["status"], "valor": tk["valor"]}).reset_index(drop=True)
    # ---- entrada na base: quanto chega DEPOIS das 09h (por dia, base e HUB) -> conversa com transporte/HUB
    e = t[no_mes(t["entrada"]) & (t["entrada"].dt.date < hoje.date())]
    entradas = (e.assign(dia=e["entrada"].dt.day, hora=e["entrada"].dt.hour + e["entrada"].dt.minute / 60)
                .groupby(["uf", "base", "hub", "dia"])
                .agg(recebidos=("waybill", "size"), apos_corte=("apos_corte", "sum"),
                     primeira=("hora", "min"), ultima=("hora", "max"), media=("hora", "mean"))
                .reset_index())
    return {"bases": bases, "emb_base": emb_base, "emb_uf": emb_uf, "diario": diario, "entradas": entradas,
            "motoristas": motoristas, "cidades": cidades, "lista": lista, "tiktok": tiktok, "hoje": hoje,
            "linhas_bd": len(df)}


def _pcts(d):
    div = lambda a, b: np.where(d[b] > 0, d[a] / d[b].replace(0, 1), np.nan)
    d["sla_a_mes"] = div("suc_mes", "dem_mes")
    d["sla_b_mes"] = div("sucb_mes", "demb_mes")
    d["sla_a_ontem"] = div("suc_ontem", "dem_ontem")
    d["sla_b_ontem"] = div("sucb_ontem", "demb_ontem")


ORDEM_PZ = {PZ_ATRASADO: 0, PZ_HOJE: 1, PZ_AMANHA: 2, PZ_SEMANA: 3}
COLS_LISTA_SLA = ["PRAZO ANJUN", "PRAZO BASE", "VALIDAÇÃO", "UF", "BASE", "Waybill", "Embarcador", "Status / Local",
                  "Motorista", "Cidade", "Dias Parado", "Valor", "Recebido na base", "Data Custódia",
                  "Data Anjun", "Data Base", "Reg. Ocorrência", "Registrante", "Motivo", "Tent.", "Rest."]


def _lista_sla(t):
    out = pd.DataFrame({
        "PRAZO ANJUN": t["pz_classe"], "PRAZO BASE": t["pzb_classe"].fillna("⚪ +7 DIAS"),
        "VALIDAÇÃO": t["validacao"], "UF": t["uf"], "BASE": t["base"], "Waybill": t["waybill"],
        "Embarcador": t["emb"], "Status / Local": t["status"], "Motorista": t["motorista"], "Cidade": t["cidade"],
        "Dias Parado": t["dias"], "Valor": t["valor"], "Recebido na base": t["entrada"], "Data Custódia": t["custodia"],
        "Data Anjun": t["prazo"], "Data Base": t["prazo_b"].dt.normalize(), "Reg. Ocorrência": t["ocorrencia"],
        "Registrante": t["registrante"], "Motivo": t["motivo"], "Tent.": t["tent"], "Rest.": t["rest"],
    })
    out["_o"] = out["PRAZO ANJUN"].map(ORDEM_PZ)
    out["_tk"] = (out["Embarcador"] != PRIORIDADE).astype(int)
    out = out.sort_values(["BASE", "_o", "_tk", "Dias Parado"], ascending=[True, True, True, False])
    out["_diff"] = t.loc[out.index, "pz_diff"]
    out["_em_rota"] = t.loc[out.index, "em_rota"]
    return out.drop(columns=["_o", "_tk"]).reset_index(drop=True)


# ================================================================ BACKLOG
COLS_LISTA_BL = ["ALERTA", "VALIDAÇÃO", "UF", "BASE", "Waybill", "Embarcador", "Status / Local", "Motorista", "Cidade",
                 "Dias Parado", "Valor", "Recebido na base", "Data Custódia", "SLA Anjun", "SLA Base",
                 "Reg. Ocorrência", "Registrante", "Motivo", "Tent.", "Rest."]
ORDEM_AL = {AL_URGENTE: 0, AL_PERDA: 1, AL_ROTA: 2, AL_ESTOQUE: 3, AL_OUTRA: 4, AL_OUTROS: 5}


def calcular_backlog(df, hoje=None):
    """
    Devolve dict:
      bases:  DataFrame uma linha por base (total, valor, aging, situação, categorias, perdas, prazos D0..D3)
      ofensores: motoristas em rota (qtd, valor) por base
      prazos_cidades: base x cidade (prazo mais comum + qtd)
      lista: todos os pacotes com a coluna ALERTA
    """
    hoje = hoje or agora()
    t = _comum(df, hoje)
    t["validacao"] = _validacao(t)
    vivo = ~t["perda"]

    # prazo da base em dias (D0, D+1, D+2, D+3+) = entrada -> SLA base; sem data: mais comum da cidade
    dd = (t["prazo_b"].dt.normalize() - t["entrada"].dt.normalize()).dt.days.clip(lower=0)
    t["d_sla"] = dd.fillna(PRAZO_BASE_PADRAO_DIAS).astype(int)

    dias = t["dias"]
    venc = t["pz_diff"] < 0
    f = pd.DataFrame({
        "uf": t["uf"], "base": t["base"],
        "total": vivo, "valor": t["valor"].where(vivo, 0.0),
        "perdas": t["perda"], "valor_perdas": t["valor"].where(t["perda"], 0.0),
        "rota": vivo & (t["categoria"] == "ROTA"), "valor_rota": t["valor"].where(vivo & (t["categoria"] == "ROTA"), 0.0),
        "estoque": vivo & (t["categoria"] == "ESTOQUE"), "valor_estoque": t["valor"].where(vivo & (t["categoria"] == "ESTOQUE"), 0.0),
        "outra_regiao": vivo & (t["categoria"] == "OUTRA REGIÃO"), "valor_outra": t["valor"].where(vivo & (t["categoria"] == "OUTRA REGIÃO"), 0.0),
        "outros": vivo & (t["categoria"] == "OUTROS"), "valor_outros": t["valor"].where(vivo & (t["categoria"] == "OUTROS"), 0.0),
        "age_mais7": vivo & (dias > 7), "age7": vivo & (dias == 7), "age6": vivo & (dias == 6), "age5": vivo & (dias == 5),
        "age4": vivo & (dias == 4), "age3": vivo & (dias == 3), "age2": vivo & (dias == 2), "age1": vivo & (dias <= 1),
        "extravio": vivo & (dias > DIAS_EXTRAVIO),
        "vencidos": vivo & (dias <= DIAS_EXTRAVIO) & venc,
        "vence_hoje": vivo & (dias <= DIAS_EXTRAVIO) & (t["pz_diff"] == 0),
        "no_prazo": vivo & (dias <= DIAS_EXTRAVIO) & ~venc & (t["pz_diff"] != 0),
        "d0": vivo & (t["d_sla"] == 0), "d1": vivo & (t["d_sla"] == 1), "d2": vivo & (t["d_sla"] == 2), "d3": vivo & (t["d_sla"] >= 3),
        "tiktok": vivo & (t["emb"] == PRIORIDADE),
    })
    bases = f.groupby(["uf", "base"]).sum(numeric_only=True).reset_index()
    def ofensor(r):
        mx = max(r["rota"], r["estoque"], r["outra_regiao"])
        if mx <= 0:
            return "-"
        return "MOTORISTA" if mx == r["rota"] else ("ESTOQUE" if mx == r["estoque"] else "OUTRA REG.")
    bases["ofensor"] = bases.apply(ofensor, axis=1) if len(bases) else []

    r = t[vivo & (t["categoria"] == "ROTA")]
    ofensores = (r.groupby(["uf", "base", "motorista"]).agg(qtd=("waybill", "size"), valor=("valor", "sum"))
                 .reset_index().sort_values(["base", "qtd"], ascending=[True, False]))

    pc = t[vivo].groupby(["uf", "base", "cidade"]).agg(
        prazo=("d_sla", lambda x: int(x.mode().iloc[0])), qtd=("waybill", "size")).reset_index()

    # ---- lista única com ALERTA
    al = pd.Series(AL_OUTROS, index=t.index, dtype="string")
    al[t["categoria"] == "OUTRA REGIÃO"] = AL_OUTRA
    al[t["categoria"] == "ESTOQUE"] = AL_ESTOQUE
    al[t["categoria"] == "ROTA"] = AL_ROTA
    al[(dias > DIAS_EXTRAVIO) | venc] = AL_URGENTE
    al[t["perda"]] = AL_PERDA
    motivo = t["motivo"].where(t["motivo"] != "", None)
    venc_dias = (-t["pz_diff"]).where(venc, 0).fillna(0).astype(int)
    motivo = motivo.where(~(t["perda"] & motivo.isna()), "Vencido há " + venc_dias.astype(str) + " dias (perda por tempo)")
    lista = pd.DataFrame({
        "ALERTA": al, "VALIDAÇÃO": t["validacao"], "UF": t["uf"], "BASE": t["base"], "Waybill": t["waybill"],
        "Embarcador": t["emb"], "Status / Local": t["status"], "Motorista": t["motorista"], "Cidade": t["cidade"],
        "Dias Parado": dias, "Valor": t["valor"], "Recebido na base": t["entrada"], "Data Custódia": t["custodia"],
        "SLA Anjun": t["prazo"], "SLA Base": t["prazo_b"], "Reg. Ocorrência": t["ocorrencia"],
        "Registrante": t["registrante"], "Motivo": motivo, "Tent.": t["tent"], "Rest.": t["rest"],
    })
    lista["_o"] = lista["ALERTA"].map(ORDEM_AL)
    lista = lista.sort_values(["BASE", "_o", "Dias Parado"], ascending=[True, True, False]).drop(columns="_o")
    return {"bases": bases, "ofensores": ofensores, "prazos_cidades": pc,
            "lista": lista.reset_index(drop=True), "hoje": hoje, "linhas_bd": len(df)}
