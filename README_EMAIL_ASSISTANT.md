# 📧 Email Assistant - Resumos Inteligentes de E-mails

Uma ferramenta que usa IA (Claude) para resumir seus e-mails do Outlook e sugerir respostas baseadas no seu estilo de comunicação.

## 🚀 O que faz

1. **Lê seus e-mails** do Outlook (dia anterior ou sexta-segunda na segunda-feira)
2. **Analisa seu histórico** de respostas anteriores para aprender seu estilo
3. **Resumo de tópicos** - Extrai os 3-5 temas principais
4. **Sugestões de resposta** - Gera respostas personalizadas para cada e-mail
5. **Envia por e-mail** - Delivers o resumo completo para sua caixa de entrada

## 📋 Pré-requisitos

- Python 3.8+
- Conta Outlook (Microsoft 365)
- Chave API do Claude (get it at https://console.anthropic.com)

## ⚙️ Setup

### 1. Instalar dependências

```bash
pip install -r requirements.txt
```

### 2. Configurar Outlook (App Password)

**Importante:** Você precisa de uma "App Password", NÃO sua senha regular.

1. Acesse: https://account.microsoft.com/account/manage-my-microsoft-account
2. Clique em **"Security"** na barra lateral
3. Vá para **"Advanced security options"**
4. Em **"App passwords"**, clique em **"Create a new app password"**
5. Copie a senha gerada (será algo como `xyzabc123def456`)

### 3. Configurar variáveis de ambiente

Crie um arquivo `.env` na pasta do projeto:

```bash
cp .env.example .env
```

Edite `.env` com suas credenciais:

```
OUTLOOK_EMAIL=seu-email@outlook.com
OUTLOOK_APP_PASSWORD=a1b2c3d4e5f6g7h8i9j0  # A senha gerada no passo 2
ANTHROPIC_API_KEY=sk-ant-...  # Sua chave API do Claude
```

## 🎯 Como usar

### Executar uma única vez

```bash
python email_assistant.py
```

Isso vai:
- ✅ Buscar seus e-mails do período
- ✅ Analisar com Claude
- ✅ Exibir o resumo no console
- ✅ Enviar por e-mail para você

### Agendar para executar automaticamente

#### Linux/Mac (com cron)

Edite seu crontab:
```bash
crontab -e
```

Adicione uma das linhas abaixo:

**Todos os dias às 8:00 AM:**
```
0 8 * * * python email_assistant.py
```

**Segunda a sexta às 9:00 AM:**
```
0 9 * * 1-5 cd /home/user/Victor-Tasks && python email_assistant.py
```

#### Windows (com Task Scheduler)

1. Abra **Task Scheduler**
2. Clique em **"Create Basic Task"**
3. Nome: "Email Summary"
4. Trigger: Escolha hora e frequência
5. Action: Programa = `C:\Python\python.exe`, Argumento = `email_assistant.py`
6. Pasta: `C:\caminho\para\Victor-Tasks`

## 📧 O que você vai receber

Um e-mail formatado com:

```
📧 Email Summary - January 10, 2025

SUMMARY OF TOPICS
- Project deadline discussed in 3 emails
- Budget review scheduled
- Team meeting feedback

SUGGESTED RESPONSES

### Email 1: From: boss@company.com - Subject: Q1 Planning
[Resposta profissional em seu estilo]

### Email 2: From: colleague@company.com - Subject: Report Feedback
[Resposta baseada em seus padrões anteriores]
```

## 🔐 Segurança

- Suas credenciais são armazenadas **apenas localmente** no arquivo `.env`
- O arquivo `.env` está no `.gitignore` (não será enviado para Git)
- Sempre use **App Password**, nunca sua senha regular
- A chave API do Claude é armazenada localmente

## 🛠️ Troubleshooting

### "Authentication failed"
- Verifique se está usando **App Password**, não sua senha regular
- Confirme que o email está correto em `.env`
- Verifique se a App Password foi copiada completamente

### "No emails found"
- Confirme que existem e-mails no período especificado
- Verifique sua conexão de internet
- Tente acessar sua caixa de entrada manualmente

### "Failed to send email"
- Confirme que a App Password está correta
- Verifique se você tem acesso à caixa de saída (Sent Items)
- Tente enviar um e-mail manualmente do Outlook

### Conexão IMAP não funciona
- Ative IMAP em sua conta Outlook:
  1. Configurações → Visualizar todas as configurações do Outlook
  2. Encaminhamento, IMAP/POP
  3. Ative IMAP

## 📝 Personalizações

### Mudar intervalo de tempo

Edite `email_assistant.py`, função `get_previous_day_range()`:

```python
# Atualmente: dia anterior (ou sexta-segunda na segunda)
# Para alterar, modifique os timedelta(days=X)
```

### Mudar idioma das análises

Edite o `system` prompt na função `analyze_emails()` para português:

```python
system="""Você é um assistente de e-mail. Sua tarefa é:
..."""
```

### Aumentar número de e-mails analisados

Na função `extract_sent_emails_for_pattern()`, aumente o parâmetro `limit`:

```python
sent_emails = self.extract_sent_emails_for_pattern(limit=50)  # Aumentado de 20
```

## 📚 Documentação

- [Anthropic Claude API](https://docs.anthropic.com)
- [Python imaplib](https://docs.python.org/3/library/imaplib.html)
- [Outlook IMAP Setup](https://support.microsoft.com/en-us/office/pop-imap-and-smtp-settings-8361e398-8af4-4e97-b9af-1474ad350e00)

## 💡 Dicas

1. **Primeiras respostas**: Os primeiros resumos serão mais genéricos. Quanto mais e-mails o sistema analisar, melhores ficarão as sugestões.

2. **Vários projetos**: Se você trabalha em múltiplos projetos, a IA agrupará automaticamente por tema.

3. **Resposta rápida**: Use as sugestões como ponto de partida - sempre revise antes de enviar!

4. **Backup de dados**: O sistema não modifica seus e-mails, apenas lê.

## 📞 Suporte

Se tiver problemas:
1. Verifique o arquivo `.env` está correto
2. Confirme que IMAP está ativado no Outlook
3. Teste a conexão IMAP manualmente com um cliente
4. Revise os logs de erro na saída do script

---

**Made with ❤️ using Claude AI**
