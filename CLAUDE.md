# CLAUDE.md — Guia pra agentes IA neste repo

Você é um Claude trabalhando no **tp-carrinhos**: app de agendamento de
**Testemunho Público com carrinhos** pra congregação JW.

Esse repo é um **spin-off do [territoryhelper](https://github.com/thiagoarce/territoryhelper)**
(2026-06). TP estava lá como MVP no commit `b86aed8` mas foi removido
porque merece domínio próprio (agenda/recorrência/notificações iam poluir
o app de território). Comece este aqui do zero, melhor.

## Stack

- **Google Apps Script** (V8) + **Google Sheets** como banco
- **Bootstrap 5** + **Leaflet** (CDN, defer)
- **clasp** pra deploy via GitHub Actions
- Tests em **Node puro** (`tests/harness.js` — sem deps)

## Status atual

Acabou de ser bootstrappado. **Nada implementado ainda**, só esqueleto:

- `appsscript.json` — runtime V8, webapp ANYONE_ANONYMOUS, USER_DEPLOYING
- `Constants.gs` — schema completo (`SHEET`, `COL`, `STATUS`, `APP_VERSION`)
- `JS_ToastPublico.html` — toast standalone (copiar do territoryhelper)
- `.github/workflows/deploy-apps-script.yml` — clasp push automático
- `tests/{harness,mocks,run}.js` — runner Node

Não tem ainda: `Code.gs`, `doGet`, `Index.html`, `Publico.html`, nenhuma
função backend. Você vai criar.

## Layout planejado

| Arquivo | Função |
|---|---|
| `Code.gs` | `doGet`, CRUDs, locks, caches, validações |
| `Utils.gs` | `withLock_`, `validarData_`, `sanitizar_`, `_dataLocalMeioDia_` |
| `Constants.gs` | Schemas (já existe — não mexe sem motivo) |
| `Index.html` | Admin (servo cadastra pontos/horários/carrinhos, vê calendário) |
| `JS_App.html` | JS principal do admin |
| `Publico.html` | Link público pro publicador agendar/check-in |
| `CSS.html` | Estilos compartilhados |

## Modelo de dados (em `Constants.gs`)

| Aba | O que guarda |
|---|---|
| `Pontos` | id, nome, lat, lng, endereco, ativo, notas |
| `Horarios` | id, pontoId, diaSemana (0-6), horaInicio, horaFim, capacidade, **recorrente**, ativo |
| `Carrinhos` | id, nome, local_guarda, ativo, notas |
| `Agendamentos` | id, horarioId, data, publicador, **pin**, **parceiroId**, carrinhoId, status, checkIn, checkOut, revistas, notas, criado |
| `Feriados` | id, data, nome, pontoId (vazio = todos os pontos) |

Status (`STATUS` enum): `agendado` → `presente` (check-in) → `concluido` (check-out).
Variantes: `ausente` (não compareceu) / `cancelado` (desmarcou).

**Schema já incorpora aprendizados do MVP** que foi removido. Não tire as colunas novas (recorrente, pin, parceiroId, aba Feriados) sem o usuário pedir.

## Convenções (herdadas do territoryhelper — siga)

### Backend (`.gs`)

- **Todo write usa `withLock_(function() { ... })`** (LockService 20s)
- **Todo write chama `_invalidar()` no fim** — limpa CacheService
- **Cache em `CacheService.getScriptCache()`** com TTL ~5min pra reads pesados
- **Invalidação versionada** (chave `*_VER` com timestamp) pra evitar percorrer chaves
- Acessos a colunas via `COL.PONTOS.X` (0-indexed) ou `COL.PONTOS.X_1IDX`
- `ensureSheet*_()` é idempotente: cria se não existe, **migra cabeçalho** se schema antigo (`if (sh.getLastColumn() < N) sh.getRange(1, ult+1, 1, ...).setValues([...])`)
- `sanitizar_(valor)` antes de gravar strings vindas do usuário
- `validarData_(yyyymmdd)` antes de aceitar datas
- **Datas em Sheets**: nunca gravar string `"yyyy-MM-dd"` direto (vira UTC midnight = dia anterior em -3). Sempre `_dataLocalMeioDia_(yyyymmdd)` que retorna Date ao meio-dia local.
- **Race em capacidade**: re-conta inscritos DENTRO do `withLock_` antes de aceitar agendamento. Frontend pode mentir.

### Frontend

- **Toast standalone** (não JS_Core do admin) — incluir via `<?!= include('JS_ToastPublico') ?>`
- **`window.toast(msg, tipo)`** em vez de `alert()` (tipos: `success` / `error` / `warn` / `info`)
- **Defer CDNs**: Bootstrap/Leaflet com `<script defer>`. Sortable/dom-to-image: lazy-load só quando precisar
- **Acumular HTML em array e fazer ONE `innerHTML`** no fim (não `innerHTML +=` em loop)
- **IDs via template server-side**: `var IDS = "<?= ids ?>";` (não `google.script.url.getLocation` — quebra silenciosamente em iOS Safari)
- **`aria-label`** em botões só-ícone
- **Min 44×44 px** em touch targets (HIG mobile)
- **Render com input do usuário**: SEMPRE escapar com `escapeHtml()` antes
- **`rel="noopener"`** em `<a target="_blank">`

### Schema migration

Cada `ensureSheet*_()` checa `getLastColumn()` e completa cabeçalho com colunas novas (idempotente, sem perda de dados). Não vai precisar disso até evoluir o schema — mas siga o padrão quando precisar.

### `_dataLocalMeioDia_(yyyymmdd)` (vai precisar logo)

```js
function _dataLocalMeioDia_(yyyymmdd) {
  if (!yyyymmdd) return '';
  var p = String(yyyymmdd).split('-');
  if (p.length !== 3) return new Date(yyyymmdd);
  return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]), 12, 0, 0);
}
```

## Roadmap (ordem de prioridade)

### Sprint 1 — Backend base
1. `Utils.gs` com `withLock_`, `validarData_`, `sanitizar_`, `_dataLocalMeioDia_`, `_invalidar()`, `limparCacheServidor()`
2. `Code.gs` com:
   - `doGet(e)` rotear `?v=admin` (Index) / `?v=publico` (Publico)
   - `include(file)`, `getScriptUrl()`
   - `ensureSheet*_()` pras 5 abas
   - CRUDs **com edit/delete**: `criar/atualizar/excluirPonto`, idem horários/carrinhos
   - `listarPontos(somenteAtivos)`, `listarHorariosDoDia(diaSemana)`, etc.
3. `agendar(payload)` com **race check** de capacidade dentro do lock
4. `checkIn(id)`, `checkOut(id, revistas)`, `cancelar(id, pin)` (PIN só dono cancela)
5. `getDadosPublico()` — pacote pra grade de agendamento

### Sprint 2 — Admin UI
6. `Index.html` + `JS_App.html` com tabs Pontos / Horários / Carrinhos / Agendamentos / Calendário
7. CRUDs visuais com edit/delete (não só create como o MVP)
8. **Calendário mensal** mostra pontos × horários × ocupação por dia (cor: cheio/parcial/vazio)
9. **Botão "Link agendamento"** copia URL `?v=publico`

### Sprint 3 — Publico UI
10. `Publico.html` com escolha de data + listagem de pontos+horários
11. **PIN** ao agendar (publicador escolhe 4 dígitos, fica em localStorage)
12. **"Minha próxima escala"** fixo no topo
13. **Check-in/out na própria UI** (com GPS opcional pra validar localização)
14. **Recorrência**: marcar "todo sábado das 8h" cria N agendamentos automáticos até data limite

### Sprint 4 — Governança
15. **Auto-ausente**: trigger diário marca quem agendou e não fez check-in
16. **Alerta de slot descoberto**: WhatsApp share / email se faltar X dias e slot vazio
17. **Swap** entre publicadores (sem cancelar/reagendar)
18. **Relatório**: horas trabalhadas, revistas distribuídas, ranking
19. **Feriados** UI: marcar Natal/Páscoa em massa

## Workflow / Deploy

- **`main`** — branch principal. Push em main = `clasp push` automático.
- Auto-deploy via `.github/workflows/deploy-apps-script.yml`. Token em `secrets.CLASP_CREDENTIALS`.
- `clasp push` envia código pra HEAD do Apps Script; `/exec` usa deployment fixo (atualizar com `clasp deploy --deploymentId $CLASP_DEPLOYMENT_ID`).
- `/dev` URL serve HEAD — bom pra dev.
- Use `feature branches` + merge pra `main` quando tests passarem.

## Setup do clasp (uma vez)

Veja `docs/clasp-setup.md` no
[territoryhelper](https://github.com/thiagoarce/territoryhelper/blob/main/docs/clasp-setup.md).
Resumo:

```bash
npm i -g @google/clasp
clasp login
# Cria projeto novo OU clona um existente
clasp create --type webapp --title "TP-Carrinhos" --rootDir .
# Edita .clasp.json com scriptId

# Pega o refresh token de ~/.clasprc.json
cat ~/.clasprc.json
# Cola em secrets.CLASP_CREDENTIALS no GitHub

# Cria deployment fixo (mantém URL /exec estável)
clasp deploy --description "Initial"
# Copia o deploymentId pra secrets.CLASP_DEPLOYMENT_ID
```

## Testes

```bash
node tests/run.js
```

Mocks de `SpreadsheetApp` em `tests/mocks.js` (sem Google real).

## Limitações conhecidas (do Apps Script)

- iframe sandboxed — sem Service Worker real
- `MailApp.sendEmail` ~100/dia em conta gratuita
- `LockService` timeout 20s no `withLock_`
- `CacheService.put` max 100KB por chave
- `google.script.url.getLocation` (postMessage) **falha silenciosamente em iOS Safari** — sempre passe params via template `<?= var ?>`

## Anti-padrões observados (não cair)

- **Regex literal dentro de template literal** — quebra o parser do HtmlService.
  Use `data-attribute` + `dataset` em vez de `${chave.replace(/'/g, "\\'")}` inline em template ``.
- **`innerHTML += html` dentro de forEach** — reparse quadrático. Sempre `array.push` e join no fim.
- **`alert()`** — bloqueia mobile, viola convenção. Sempre `window.toast()`.
- **Date a partir de `"yyyy-MM-dd"` direto** — sempre `_dataLocalMeioDia_`.

## Quando o usuário disser "manda bala"

Faça um commit isolado por feature. Cada commit testável independentemente.
Deploy a cada commit (push em `main` dispara workflow). Se quebrar,
`git revert` isolado.

Boa sorte 🚀
