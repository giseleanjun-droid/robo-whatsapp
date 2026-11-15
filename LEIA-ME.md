# Robô do WhatsApp (repositório público `robo-whatsapp`)

Às **08:47, todos os dias** o GitHub liga um computador na nuvem, lê a aba **📲 WHATSAPP** da planilha
**MATRIZ SLA + BACKLOG** (feita pelo robô das 04h) e manda a mensagem de cada base no grupo dela.
Notebook pode ficar desligado. Repositório público = minutos ilimitados e grátis.

**O que é público:** só este código. **O que NÃO é público:** as mensagens (ficam no seu Drive), a sessão do
WhatsApp (guardada criptografada), o número e os e-mails (segredos do GitHub). O log não mostra mensagens,
nomes de grupos nem o código de pareamento.

## Segredos (github.com > robo-whatsapp > Settings > Secrets and variables > Actions)

| Segredo | O que é |
|---|---|
| `DRIVE_TOKEN_JSON` | o mesmo JSON da conta de serviço do robô (`NUVEM\github\conta_servico.json`) |
| `PASTA_SLA` | id da pasta "Bases Consolidadas" do SLA |
| `WA_TELEFONE` | número do WhatsApp com DDI e DDD, só números (ex.: 5562999998888) |
| `WA_CHAVE` | uma senha longa qualquer: tranca a sessão do WhatsApp. Não precisa decorar |
| `ALERTA_REMETENTE`, `ALERTA_SENHA`, `ALERTA_PARA` | os mesmos do robô (e-mail Gmail + senha de app) |

## Conectar o WhatsApp (uma vez, ou quando cair)

1. GitHub > robo-whatsapp > **Actions** > **WhatsApp Anjun** > **Run workflow** > modo = `parear`.
2. Chega um **e-mail com um código** de 8 letras/números.
3. No celular: **WhatsApp > Dispositivos conectados > Conectar dispositivo > Conectar com número de telefone**
   e digite o código (você tem uns 4 minutos).
4. Pronto: o robô aparece como mais um "aparelho conectado" (como o WhatsApp Web).

## Dizer qual é o grupo de cada base

Na planilha **GRUPOS WHATSAPP** (pasta do SLA), coluna **GRUPO WHATSAPP**: o nome do grupo **exatamente** como
aparece no WhatsApp (acento e maiúscula não importam). Ou cole o **link de convite** do grupo na coluna
LINK DO GRUPO (mais seguro se dois grupos tiverem nomes parecidos). Base sem grupo = não envia.
**Para não enviar a uma base (inativa, ou um estado inteiro, ex.: MS): apague a linha dela** -- o robô não devolve.
O número conectado precisa estar **dentro** do grupo.

## Testar sem incomodar ninguém

Crie um grupo só seu, grave o link de convite dele no segredo `WA_GRUPO_TESTE` e rode o workflow com `teste = sim`:
todas as mensagens vão só para esse grupo. Para testar poucas, preencha `so_bases` (ex.: `GO-W-D001`).
Modo `grupos` só confere quais bases acharam o grupo, sem enviar nada.

## Todo dia

- Chega um e-mail: quantas foram enviadas e quais não (sem grupo / grupo não encontrado).
- Se a matriz não for de hoje (robô das 04h falhou), **não envia** mensagem velha e avisa por e-mail.
- Se a sessão cair: e-mail "precisa parear de novo" -> repita "Conectar o WhatsApp".

## Cuidados

Automação de WhatsApp não é oficial. Para diminuir o risco: uma mensagem por grupo por dia, com pausa de
8 a 15 s entre os grupos e "digitando..." antes de cada envio. O servidor do GitHub fica fora do Brasil, então
o WhatsApp pode derrubar a sessão de vez em quando (aí é só parear de novo).
