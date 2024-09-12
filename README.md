# MailMe

MailMe is a project written in TypeScript that enables sending emails via Microsoft Graph, with or without attachments. The goal is to simplify email sending with Outlook using a straightforward and intuitive command line.

## Running this project

For the application to function correctly, certain parameters are required. These should be specified in the appSettings.ts file. The information needed includes clientId, client_secret, and directory_id. To obtain these credentials for using Microsoft Graph, follow the step-by-step guide available at https://wiki.hexing.com.br/books/glpi/page/azure-microsoft-graph-oauth2-imap-e-mails. Once you have completed this step, you will be ready to fill in appSettings, where:

clientId <=> Application (client) ID
client_secret <=> Client secret
directory_id <=> Directory (tenant) ID

### Using PM2 to start

To start the application using PM2, run the command below in the directory containing the index.ts file.

```bash
pm2 start index.ts --interpreter ts-node --watch --ignore-watch="uploads teste_de_envio"
```

By default, the application will run on port 3000. Use proxy reverse to access externally.

### Log in to get a token

```bash
curl -X POST https://url.com/mailme/login -F 'data={"user":"user", "password" :"password"}'
```

The response will be a JSON containing the token. Use this token in subsequent requests.

It is important to read the logs using the command below.

```bash
pm2 logs
```

## Features

To send an email with an attachment, use the structure below as a reference:

```bash
curl -X POST https://url.com/mailme/ -F "file[]=@/caminho/do/arquivo/arquivo_1.txt" -F "file[]=@/caminho/do/arquivo/arquivo_2.txt"  -F "data={\"subject\":\"Assunto\",\"body\":\"Corpo da mensagem\",\"sendto\":\"username@domain.com\"}" -H "authorization: token"
```

To send an email without an attachment, only with a subject and body, use the structure below as a reference:

```bash
curl -X POST https://url.com/mailme/  -F "data={\"subject\":\"Assunto\",\"body\":\"Corpo da mensagem\",\"sendto\":\"username@domain.com\"}"   -H "authorization: token"
```
