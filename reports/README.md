# Relatório FinançasCasal via Telegram

Envia um resumo financeiro (saldos, gastos, faturas a vencer, parcelas, orçamento,
patrimônio) pro Telegram de vocês dois, automático, pelo GitHub Actions.

- **Toda segunda 08:00** → resumo da semana / mês parcial
- **Todo dia 1º 08:00** → fechamento do mês anterior
- Botão **Run workflow** (aba Actions) → dispara na hora, pra testar

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
| `TELEGRAM_CHAT_IDS` | os dois chat ids separados por vírgula, ex.: `123456789,987654321` |

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
