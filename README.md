# TP-Carrinhos

App de agendamento de **Testemunho Público com carrinhos** pra congregação.

Spun-off do [territoryhelper](https://github.com/thiagoarce/territoryhelper)
em 2026-06 quando ficou claro que TP merece domínio próprio: recorrência,
check-in/out, swap, calendário, notificações são features que iam poluir
o app de território.

## Reuso do territoryhelper

Já neste bootstrap:
- `JS_ToastPublico.html` — toast standalone, sem deps
- `appsscript.json` — runtime V8 + webapp ANYONE_ANONYMOUS + USER_DEPLOYING
- `.claspignore` — só envia `.gs`/`.html`/`appsscript.json` pro Apps Script
- `.github/workflows/deploy-apps-script.yml` — clasp push automático em main
- `tests/harness.js` + `tests/mocks.js` — sheets fake em Node, sem Google
- `Constants.gs` — schema próprio (já com melhorias do MVP)

## Schema já com aprendizados do MVP

Diferenças desde o MVP que estava no territoryhelper:

- **`recorrente`** flag em horários (gera N agendamentos automáticos)
- **`parceiroId`** em agendamentos (pareamento explícito de irmãos)
- **`pin`** em agendamentos (proteção contra outro cancelar)
- aba **`Feriados`** (cancela datas em massa por ponto ou geral)

## A fazer (em ordem de prioridade)

1. **CRUDs com edit/delete** (no MVP só tinha create)
2. **Race check de capacidade** dentro de `withLock_` (no MVP só frontend)
3. **Auto-ausente** — cron diário marca quem não fez check-in
4. **Check-in/out na UI** pública (backend tinha mas tela não usava)
5. **Calendário mensal** pro servo ver buracos
6. **Visão "minha próxima escala"** pro publicador
7. **Notificação WhatsApp/email** de slot descoberto
8. **Swap entre publicadores** (não cancelar+reagendar)
9. **Relatório** de horas + revistas distribuídas

## Setup

Veja o setup do clasp/workflow no
[territoryhelper docs/clasp-setup.md](https://github.com/thiagoarce/territoryhelper/blob/main/docs/clasp-setup.md).

```bash
npm i -g @google/clasp
clasp login
clasp create --type webapp --title "TP-Carrinhos" --rootDir .
# Cria deployment e copia o ID pra secret CLASP_DEPLOYMENT_ID
# Cria secret CLASP_CREDENTIALS no GitHub (refresh_token do clasp login)
```

## Testes

```bash
node tests/run.js
```
