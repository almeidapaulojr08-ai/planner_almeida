# Relatório FinançasCasal via Telegram

Envia um resumo financeiro (saldos, gastos, faturas a vencer, parcelas, orçamento,
patrimônio) pro Telegram de vocês dois, automático, pelo GitHub Actions.

- **Toda segunda-feira às 08:07** (Brasília) → cada um recebe 2 mensagens: uma "Geral (casal)" e outra só com os valores dele. (Era 2x ao dia até 23/09/2026; Paulo pediu semanal.)
- No **dia 1º** o texto vira "fechamento do mês anterior".
- Botão **Run workflow** (aba Actions) → dispara na hora, pra testar (ignora a trava de 5 h).
- **Backup semanal:** no mesmo run, `send_backup.js` manda o JSON completo do banco como arquivo pros dois chats (mesmo formato do Exportar JSON; restaura pelo Importar JSON). O backup sai SEMPRE que o workflow roda, mesmo quando o relatório é pulado pela trava.
- Trava anti-duplicata: se já enviou há menos de 5 h, o script pula (grava `reportMeta/lastSentAt` no Firebase).
- Trocar horário: editar o `cron` em `.github/workflows/relatorio.yml` (está em UTC; BRT = UTC-3) **e** o horário no cron-job.org (abaixo).

### ✅ Gatilho em uso: tarefa agendada no PC do Paulo (23/09/2026)
Tarefa do Windows **"FinancasCasal - Relatorio Telegram"** (Agendador de Tarefas) roda `reports/disparar_relatorio.cmd`
toda **segunda-feira às 08:07**, que chama `gh workflow run relatorio.yml -f force=false`. Se o PC estava desligado no horário,
roda assim que ligar. Log em `%LOCALAPPDATA%incasal_relatorio.log`. Se um dia quiser algo que não dependa do PC,
usar o cron-job.org (abaixo).

### ⚠️ O cron do GitHub NÃO é confiável
Em 23/09/2026 os dois agendamentos do dia (08:00 e 20:00) simplesmente não dispararam, sem erro nenhum —
o GitHub trata `schedule` como "melhor esforço" e pula quando está sobrecarregado. Por isso o gatilho
principal é o **cron-job.org** (gratuito, pontual), que chama a API do GitHub e dispara o workflow.
O cron do GitHub fica só de backup; a trava de 5 h evita mensagem em dobro quando os dois funcionam.

#### Configurar o cron-job.org (uma vez)
1. **Token do GitHub:** GitHub → foto de perfil → *Settings* → *Developer settings* → *Personal access tokens* →
   *Fine-grained tokens* → *Generate new token*.
   - Repository access: *Only select repositories* → `planner_almeida`
   - Permissions → Repository permissions → **Actions: Read and write**
   - Expiration: 1 ano (anota a data pra renovar). Copia o token (`github_pat_...`).
2. **Conta no cron-job.org:** https://cron-job.org → *Sign up* (grátis).
3. **Criar o job** (*Create cronjob*):
   - Title: `Relatório FinançasCasal`
   - URL: `https://api.github.com/repos/almeidapaulojr08-ai/planner_almeida/actions/workflows/relatorio.yml/dispatches`
   - Schedule → *Custom* → Timezone `America/Sao_Paulo`, horas `8,20`, minuto `7`
   - Aba **Advanced**:
     - Request method: **POST**
     - Headers: `Authorization` = `Bearer github_pat_...` · `Accept` = `application/vnd.github+json` · `Content-Type` = `application/json`
     - Request body: `{"ref":"main","inputs":{"force":"false"}}`
   - Salvar. Botão *Test run* → tem que voltar **HTTP 204** (sem corpo). Aí olha a aba Actions: apareceu um run novo.
4. Se um dia o token expirar, o cron-job.org passa a receber 401 e manda e-mail avisando — é só gerar outro token e trocar no header.

---

## Setup (uma vez só) — checklist

### 1. Criar o bot do Telegram
1. No Telegram, fala com **@BotFather** → `/newbot`
2. Dá um nome e um usuário pro bot (ex.: `financas_casal_bot`)
3. Ele te devolve um **token** tipo `8123456789:AAH...` → guarda

### 2. Descobrir os chat IDs (Paulo e Thayse)
1. Cada um abre o Telegram, procura o bot que você criou e manda um **"oi"** (isso "autoriza" o bot a te mandar mensagem)
2. No navegador, abre (troca `<TOKEN>` pelo do bot):
   `https://api.telegram.org/bot<TOKEN>/getUpdates`
3. Procura `"chat":{"id":123456789` — esse número é o chat id. Pega o de cada um.
   - Se aparecer vazio, manda "oi" pro bot de novo e recarrega a página.

### 3. Gerar a chave de acesso ao Firebase (service account)
1. Console do Firebase → ⚙️ **Configurações do projeto** → aba **Contas de serviço**
2. **Gerar nova chave privada** → baixa um arquivo `.json`
3. Abre o `.json`, copia **todo o conteúdo** (é o que vai no secret abaixo)
   - ⚠️ Esse arquivo dá acesso total ao banco. **NUNCA** commita ele no repositório — só cola como secret.

### 4. Cadastrar os secrets no GitHub
No repositório `planner_almeida` → **Settings** → **Secrets and variables** → **Actions** → **New repository secret**. Cria estes quatro:

| Nome do secret | Valor |
|---|---|
| `FIREBASE_SERVICE_ACCOUNT` | cole o conteúdo inteiro do `.json` do passo 3 |
| `FIREBASE_DB_URL` | `https://almeida-wosniak-dre-default-rtdb.firebaseio.com` |
| `TELEGRAM_BOT_TOKEN` | o token do passo 1 |
| `TELEGRAM_CHAT_IDS` | pares `id:Nome` separados por vírgula, ex.: `111:Paulo,222:Thayse`. O Nome tem que ser igual ao das transações. |
| `ANTHROPIC_API_KEY` | **(opcional)** chave da Anthropic (`sk-ant-...`). Com ela, a mensagem Geral ganha 2–3 frases de leitura escritas pelo Claude ("setembro está acima da média, puxado por Casa..."). Custa centavos por dia. Sem a secret, o relatório sai igual ao de sempre. |

### 5. Testar
- Aba **Actions** → workflow **Relatório FinançasCasal** → **Run workflow** → aguarda ~1 min
- Se tudo certo, chega a mensagem no Telegram de vocês.
- Deu erro? Abre o run, olha o log do passo "Enviar relatório" — a mensagem de erro diz o que faltou (secret errado, chat id sem "oi" pro bot, etc.)

---

## Rodar localmente (opcional, pra testar sem o Actions)
```bash
cd reports
npm install
export FIREBASE_SERVICE_ACCOUNT="$(cat caminho/para/serviceAccount.json)"
export FIREBASE_DB_URL="https://almeida-wosniak-dre-default-rtdb.firebaseio.com"
export TELEGRAM_BOT_TOKEN="..."
export TELEGRAM_CHAT_IDS="123,456"
node send_report.js
```

## WhatsApp depois
A lógica do relatório (o `send_report.js`) fica igual. Pra trocar Telegram por
WhatsApp oficial (Meta Cloud API), muda só o trecho final que faz o `fetch` de
envio — o cálculo do resumo é reaproveitado 100%.
