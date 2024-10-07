# enviar-email-msgraph

Biblioteca para enviar e-mails usando o Microsoft Graph API em Node.js.

## Descrição

Esta biblioteca permite enviar e-mails através do Microsoft Graph API. Ela encapsula a lógica de obtenção do token de acesso e o envio de e-mails, simplificando a integração com o Microsoft Graph em projetos Node.js.

## Instalação

Para instalar este pacote, use o npm:

```bash
npm install enviar-email-msgraph
```

## Como Usar

### 1. Importar o Pacote

Primeiro, importe o pacote e crie uma nova instância de `GraphEmailClient`, fornecendo as credenciais necessárias:

```javascript
const GraphEmailClient = require("enviar-email-msgraph");

// Crie uma instância do cliente de e-mail com as credenciais
const emailClient = new GraphEmailClient(
  "YOUR_CLIENT_ID",
  "YOUR_CLIENT_SECRET",
  "YOUR_TENANT_ID",
  "YOUR_USER_ID"
);
```

### 2. Enviar um E-mail

Use o método `sendEmail` da instância para enviar um e-mail:

```javascript
emailClient
  .sendEmail("destinatario@exemplo.com", "Assunto do E-mail", "Corpo do E-mail")
  .then(() => console.log("E-mail enviado com sucesso!"))
  .catch((error) => console.error("Erro ao enviar e-mail:", error));
```

### Parâmetros

- **`clientId`**: O ID do cliente registrado no Azure AD.
- **`clientSecret`**: O segredo do cliente gerado no Azure AD.
- **`tenantId`**: O ID do locatário associado ao Azure AD.
- **`userId`**: O endereço de e-mail do usuário que enviará os e-mails (deve ter permissões adequadas).
- **`toEmailAddress`**: O endereço de e-mail do destinatário.
- **`subject`**: O assunto do e-mail.
- **`bodyContent`**: O conteúdo do corpo do e-mail.

## Exemplo Completo

Aqui está um exemplo completo de como configurar e enviar um e-mail:

```javascript
const GraphEmailClient = require("enviar-email-msgraph");

// Configuração do cliente de e-mail
const emailClient = new GraphEmailClient(
  "YOUR_CLIENT_ID",
  "YOUR_CLIENT_SECRET",
  "YOUR_TENANT_ID",
  "YOUR_USER_ID"
);

// Enviando o e-mail
emailClient
  .sendEmail("destinatario@exemplo.com", "Assunto do E-mail", "Corpo do E-mail")
  .then(() => console.log("E-mail enviado com sucesso!"))
  .catch((error) => console.error("Erro ao enviar e-mail:", error));
```

## Dependências

Este pacote depende de:

- `node-fetch`: Para fazer solicitações HTTP ao Microsoft Graph API.

## Observações

1. **Permissões**: O aplicativo registrado no Azure AD deve ter as permissões corretas para enviar e-mails através do Microsoft Graph API (`Mail.Send`).
2. **Segurança**: Nunca inclua suas credenciais (`clientId`, `clientSecret`, etc.) diretamente no código-fonte em ambientes de produção. Utilize variáveis de ambiente ou serviços seguros para gerenciar essas informações.

## Licença

ISC License


### Resumo

Este `README.md` inclui:
- Uma breve descrição do pacote.
- Instruções para instalação.
- Um exemplo de uso com explicações detalhadas sobre como configurar e enviar um e-mail.
- Descrição dos parâmetros e métodos.
- Observações importantes sobre permissões e segurança.