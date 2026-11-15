// enviar.mjs - manda a mensagem de cada base no grupo de WhatsApp dela.
//
// MODO=parear  -> pede um "código de pareamento" para o número em WA_TELEFONE e grava em codigo.txt
//                 (a Gisele digita o código no celular: WhatsApp > Dispositivos conectados >
//                 Conectar dispositivo > Conectar com número de telefone). A sessão fica em auth/.
// MODO=enviar  -> lê mensagens.json e envia. Pausa aleatória entre os grupos (parecer gente).
// MODO=grupos  -> só lista quantos grupos a conta tem e quais bases acharam grupo (não envia nada).
// GRUPO_TESTE  -> se preenchido (nome do grupo), TODAS as mensagens vão para esse grupo (teste).
//
// O log é PÚBLICO (repositório público): aqui só aparecem códigos de base e contagens,
// nunca o texto das mensagens, nomes de grupos, números ou o código de pareamento.
import fs from "node:fs";
import { execFileSync } from "node:child_process";
import pino from "pino";
import * as B from "@whiskeysockets/baileys";

const makeWASocket = B.makeWASocket || B.default?.default || B.default;
const { useMultiFileAuthState, DisconnectReason, fetchLatestBaileysVersion, Browsers } = B;

const MODO = (process.env.MODO || "enviar").toLowerCase();
const AUTH = process.env.AUTH_DIR || "auth";
const PAUSA_MIN = +(process.env.PAUSA_MIN || 8), PAUSA_MAX = +(process.env.PAUSA_MAX || 15);
const GRUPO_TESTE = (process.env.GRUPO_TESTE || "").trim();
const espera = (ms) => new Promise((r) => setTimeout(r, ms));
const norm = (s) => String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/\s+/g, " ").trim().toUpperCase();
const log = (...a) => console.log(new Date().toLocaleTimeString("pt-BR", { timeZone: "America/Sao_Paulo" }), ...a);

async function conectar() {
  const { state, saveCreds } = await useMultiFileAuthState(AUTH);
  let version;
  try { ({ version } = await fetchLatestBaileysVersion()); } catch { /* usa a padrão da biblioteca */ }
  const sock = makeWASocket({
    version, auth: state, logger: pino({ level: "silent" }), browser: Browsers.ubuntu("Chrome"),
    markOnlineOnConnect: false, syncFullHistory: false, printQRInTerminal: false,
  });
  sock.ev.on("creds.update", saveCreds);
  return { sock, state };
}

// abre a conexão; reconecta 1x quando o WhatsApp pede "restart" (acontece logo depois do pareamento)
async function abrir({ parear = false, prazoMs = 240000 } = {}) {
  const fim = Date.now() + prazoMs;
  for (let tentativa = 1; tentativa <= 4 && Date.now() < fim; tentativa++) {
    // pareando e ainda não registrado: começa sempre do zero (sessão pela metade faz o WhatsApp recusar)
    if (parear && fs.existsSync(`${AUTH}/creds.json`)) {
      const c = JSON.parse(fs.readFileSync(`${AUTH}/creds.json`, "utf-8"));
      if (!c.me) fs.rmSync(AUTH, { recursive: true, force: true });
    }
    const { sock, state } = await conectar();
    let pediuCodigo = false;
    const r = await new Promise((resolve) => {
      const t = setTimeout(() => resolve({ ok: false, motivo: "tempo esgotado" }), Math.max(1000, fim - Date.now()));
      sock.ev.on("connection.update", async (u) => {
        if (u.qr && parear && !pediuCodigo && !state.creds.me) {
          pediuCodigo = true;
          const tel = String(process.env.WA_TELEFONE || "").replace(/\D/g, "");
          if (!tel) { clearTimeout(t); return resolve({ ok: false, motivo: "WA_TELEFONE vazio" }); }
          try {
            const codigo = await sock.requestPairingCode(tel);
            fs.writeFileSync("codigo.txt", codigo);
            // e-mail NA HORA (antes: só saía quando o robô desistia de esperar)
            try { execFileSync("python", ["relatorio.py", "codigo"], { stdio: "ignore", timeout: 60000 }); }
            catch { log("falhou o e-mail do código"); }
            log(`código ${tentativa} gerado e enviado por e-mail. Aguardando você digitar no celular...`);
          } catch (e) { clearTimeout(t); return resolve({ ok: false, motivo: "não gerou o código: " + e.message }); }
        } else if (u.qr && !parear) {
          clearTimeout(t); return resolve({ ok: false, deslogado: true, motivo: "sessão não existe ou caiu (precisa parear)" });
        }
        if (u.connection === "open") { clearTimeout(t); resolve({ ok: true, sock }); }
        if (u.connection === "close") {
          const cod = u.lastDisconnect?.error?.output?.statusCode;
          clearTimeout(t);
          // pareando: código venceu (408) ou foi recusado -> tenta de novo com código novo
          if (parear && !state.creds.me) return resolve({ ok: false, reconectar: true, motivo: "código venceu, gerando outro (" + cod + ")" });
          if (cod === DisconnectReason.loggedOut) return resolve({ ok: false, deslogado: true, motivo: "o WhatsApp desconectou este aparelho" });
          resolve({ ok: false, reconectar: true, motivo: "conexão fechada (" + cod + ")" });
        }
      });
    });
    if (r.ok) return r.sock;
    if (!r.reconectar) return Promise.reject(Object.assign(new Error(r.motivo), { deslogado: !!r.deslogado }));
    log("reconectando:", r.motivo);
    sock.end?.(undefined);
    await espera(3000);
  }
  throw new Error("não conectou no tempo");
}

async function mapaGrupos(sock) {
  const todos = await sock.groupFetchAllParticipating();
  const porNome = new Map();
  for (const [jid, g] of Object.entries(todos)) porNome.set(norm(g.subject), jid);
  return { porNome, total: Object.keys(todos).length };
}

let jidTeste;
async function porLink(sock, link) {
  const cod = String(link || "").trim().split("?")[0].split("/").pop();
  if (!cod || cod.length < 10) return undefined;
  try { const info = await sock.groupGetInviteInfo(cod); return info?.id; } catch { return undefined; }
}

async function jidDoGrupo(sock, item, porNome) {
  if (GRUPO_TESTE) {   // aceita o link de convite ou o nome do grupo de teste
    if (jidTeste === undefined) jidTeste = (await porLink(sock, GRUPO_TESTE)) || porNome.get(norm(GRUPO_TESTE)) || null;
    return jidTeste || undefined;
  }
  const cod = (item.link || "").trim().split("?")[0].split("/").pop();
  if (cod && cod.length >= 10) {
    try { const info = await sock.groupGetInviteInfo(cod); if (info?.id) return info.id; } catch { /* tenta pelo nome */ }
  }
  return item.grupo ? porNome.get(norm(item.grupo)) : undefined;
}

async function main() {
  if (MODO === "parear") {
    fs.rmSync(AUTH, { recursive: true, force: true });   // pareamento sempre começa do zero
    const sock = await abrir({ parear: true, prazoMs: 600000 });
    log("PAREADO: sessão salva.");
    await espera(5000);   // deixa gravar as chaves
    sock.end?.(undefined);
    return 0;
  }

  const sock = await abrir({ prazoMs: 120000 });
  log("conectado ao WhatsApp");
  await espera(4000);
  const { porNome, total } = await mapaGrupos(sock);
  const itens = JSON.parse(fs.readFileSync("mensagens.json", "utf-8"));
  log(`grupos na conta: ${total} · bases na matriz: ${itens.length}${GRUPO_TESTE ? " · MODO TESTE (tudo no grupo de teste)" : ""}`);

  const resultado = [];
  for (const it of itens) {
    if (!GRUPO_TESTE && !it.grupo && !it.link) { resultado.push({ base: it.base, status: "sem grupo configurado" }); continue; }
    const jid = await jidDoGrupo(sock, it, porNome);
    if (!jid) { resultado.push({ base: it.base, status: "grupo não encontrado" }); log(it.base, "grupo não encontrado"); continue; }
    if (MODO === "grupos") { resultado.push({ base: it.base, status: "grupo ok (não enviado)" }); continue; }
    try {
      await sock.presenceSubscribe(jid).catch(() => {});
      await sock.sendPresenceUpdate("composing", jid).catch(() => {});
      await espera(1500 + Math.random() * 2000);
      await sock.sendMessage(jid, { text: it.texto });
      await sock.sendPresenceUpdate("paused", jid).catch(() => {});
      resultado.push({ base: it.base, status: "enviado", hora: new Date().toISOString() });
      log(it.base, "enviado");
    } catch (e) {
      resultado.push({ base: it.base, status: "falhou: " + e.message });
      log(it.base, "falhou");
    }
    await espera((PAUSA_MIN + Math.random() * (PAUSA_MAX - PAUSA_MIN)) * 1000);
  }
  fs.writeFileSync("resultado.json", JSON.stringify(resultado, null, 1));
  const env = resultado.filter((r) => r.status === "enviado").length;
  // marca o dia como enviado (as outras tentativas agendadas da manhã não mandam de novo)
  if (MODO === "enviar" && !GRUPO_TESTE && env > 0) fs.writeFileSync("enviado.txt", new Date().toISOString());
  log(`FIM: ${env} enviadas · ${resultado.length - env} não enviadas`);
  await espera(3000);
  sock.end?.(undefined);
  return resultado.some((r) => r.status.startsWith("falhou") || r.status === "grupo não encontrado") ? 1 : 0;
}

main().then((c) => process.exit(c)).catch((e) => {
  log("ERRO:", e.message);
  fs.writeFileSync("resultado.json", JSON.stringify([{ base: "-", status: (e.deslogado ? "DESLOGADO: " : "ERRO: ") + e.message }]));
  process.exit(e.deslogado ? 2 : 1);
});
