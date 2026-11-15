# -*- coding: utf-8 -*-
"""
relatorio.py - manda para a Gisele (e-mail Gmail com senha de app, os mesmos segredos do robô):
  python relatorio.py codigo   -> o código de pareamento (codigo.txt)
  python relatorio.py envio    -> resumo do envio (resultado.json / motivo.txt)
Nada disso aparece no log público.
"""
import json
import os
import smtplib
import ssl
import sys
import datetime as dt
from email.message import EmailMessage
from pathlib import Path

AQUI = Path(__file__).resolve().parent


def mandar(assunto, corpo):
    rem, senha, para = (os.environ.get(k, "").strip() for k in ("ALERTA_REMETENTE", "ALERTA_SENHA", "ALERTA_PARA"))
    if not (rem and senha and para):
        print("(e-mail não configurado)")
        return
    m = EmailMessage()
    m["From"], m["To"], m["Subject"] = rem, para, assunto
    m.set_content(corpo)
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=ssl.create_default_context()) as s:
        s.login(rem, senha)
        s.send_message(m)
    print("e-mail enviado")


def ler(nome):
    p = AQUI / nome
    return p.read_text(encoding="utf-8") if p.exists() else ""


def main():
    tipo = sys.argv[1] if len(sys.argv) > 1 else "envio"
    agora = dt.datetime.now().strftime("%d/%m %H:%M")
    if tipo == "codigo":
        cod = ler("codigo.txt").strip()
        if not cod:
            return
        mandar(f"WhatsApp robô: código de pareamento {cod}",
               f"Código: {cod}\n\nNo celular do número da Anjun, em até 4 minutos:\n"
               "WhatsApp > ⋮ (ou Configurações) > Dispositivos conectados > Conectar dispositivo >\n"
               "'Conectar com número de telefone' e digite o código acima.\n\n"
               "Se o tempo passar, rode o workflow de novo com modo = parear.")
        return

    motivo = ler("motivo.txt").strip()
    if motivo:
        mandar("WhatsApp robô: mensagens NÃO enviadas hoje",
               f"{agora} - {motivo}\n\nNenhuma mensagem foi enviada. Confira se o robô das 04h rodou "
               "(github.com > robo-anjun > Actions) e, depois, rode o robo-whatsapp de novo (Run workflow).")
        return
    try:
        res = json.loads(ler("resultado.json") or "[]")
    except Exception:
        res = []
    if not res:
        mandar("WhatsApp robô: não rodou até o fim", f"{agora} - veja o log em github.com > robo-whatsapp > Actions.")
        return
    if res[0].get("status", "").startswith("DESLOGADO"):
        mandar("WhatsApp robô: a sessão do WhatsApp caiu - precisa parear de novo",
               f"{agora} - {res[0]['status']}\n\nNo GitHub: robo-whatsapp > Actions > WhatsApp Anjun > Run workflow,\n"
               "modo = parear. Você recebe outro e-mail com o código para digitar no celular.")
        return
    ok = [r for r in res if r["status"] == "enviado"]
    nao = [r for r in res if r["status"] != "enviado"]
    corpo = f"{agora} - {len(ok)} mensagem(ns) enviada(s), {len(nao)} não enviada(s).\n\n"
    if nao:
        corpo += "NÃO ENVIADAS:\n" + "\n".join(f"- {r['base']}: {r['status']}" for r in nao) + "\n\n"
        corpo += "Bases 'sem grupo' ou 'grupo não encontrado': preencha o nome exato do grupo (ou o link de convite)\n" \
                 "na planilha GRUPOS WHATSAPP. Para NÃO enviar a uma base, apague a linha dela nessa planilha.\n\n"
    if ok:
        corpo += "ENVIADAS:\n" + ", ".join(r["base"] for r in ok)
    mandar(f"WhatsApp robô: {len(ok)} enviadas, {len(nao)} pendentes", corpo)


if __name__ == "__main__":
    main()
