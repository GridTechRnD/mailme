// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

// <UserAuthConfigSnippet>
import 'isomorphic-fetch';
import { DeviceCodeCredential, DeviceCodePromptCallback } from '@azure/identity';
import { Client, PageCollection } from '@microsoft/microsoft-graph-client';
import { User, Message } from '@microsoft/microsoft-graph-types';
import { TokenCredentialAuthenticationProvider } from
  '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';
import { AppSettings } from './appSettings';
import * as fs from 'fs';

let _settings: AppSettings | undefined = undefined;
let _deviceCodeCredential: DeviceCodeCredential | undefined = undefined;
let _userClient: Client | undefined = undefined;

export function initializeGraphForUserAuth(settings: AppSettings, deviceCodePrompt: DeviceCodePromptCallback) {
  // Ensure settings isn't null
  if (!settings) {
    throw new Error('Settings cannot be undefined');
  }

  _settings = settings;

  _deviceCodeCredential = new DeviceCodeCredential({
    clientId: settings.clientId,
    tenantId: settings.tenantId,
    userPromptCallback: deviceCodePrompt
  });

  const authProvider = new TokenCredentialAuthenticationProvider(_deviceCodeCredential, {
    scopes: settings.graphUserScopes
  });

  _userClient = Client.initWithMiddleware({
    authProvider: authProvider
  });
}
// </UserAuthConfigSnippet>

// <GetUserTokenSnippet>
export async function getUserTokenAsync(): Promise<string> {
  // Ensure credential isn't undefined
  if (!_deviceCodeCredential) {
    throw new Error('Graph has not been initialized for user auth');
  }

  // Ensure scopes isn't undefined
  if (!_settings?.graphUserScopes) {
    throw new Error('Setting "scopes" cannot be undefined');
  }

  // Request token with given scopes
  const response = await _deviceCodeCredential.getToken(_settings?.graphUserScopes);
  return response.token;
}
// </GetUserTokenSnippet>

// <GetUserSnippet>
export async function getUserAsync(): Promise<User> {
  // Ensure client isn't undefined
  if (!_userClient) {
    throw new Error('Graph has not been initialized for user auth');
  }

  return _userClient.api('/me')
    // Only request specific properties
    .select(['displayName', 'mail', 'userPrincipalName'])
    .get();
}
// </GetUserSnippet>

// <GetInboxSnippet>
export async function getInboxAsync(): Promise<PageCollection> {
  // Ensure client isn't undefined
  if (!_userClient) {
    throw new Error('Graph has not been initialized for user auth');
  }

  return _userClient.api('/me/mailFolders/inbox/messages')
    .select(['from', 'isRead', 'receivedDateTime', 'subject'])
    .top(25)
    .orderby('receivedDateTime DESC')
    .get();
}
// </GetInboxSnippet>

// <SendMailSnippet>
export async function sendMailAsync(subject: string, body: string, recipient: string) {
  // Ensure client isn't undefined
  if (!_userClient) {
    throw new Error('Graph has not been initialized for user auth');
  }

  // Create a new message
  const message: Message = {
    subject: subject,
    body: {
      contentType: 'html',
      content: body
    },
    toRecipients: [
      {
        emailAddress: {
          address: recipient
        }
      }
    ]
  };

  // Send the message
  return _userClient.api('me/sendMail')
    .post({
      message: message
    });
}
// </SendMailSnippet>

// Função para obter a assinatura do usuário
async function getUserSignature(): Promise<string> {
  try {
      // Ensure client isn't undefined
      if (!_userClient) {
        throw new Error('Graph has not been initialized for user auth');
      }
      const response = await _userClient.api('/me/mailFolders/AAMkAGVhMzRjNzk0LTJlMmMtNDUzZS05NjY3LWYzODBhMDRiZTFhOAAuAAAAAAA2iaXIinf6RZfwk5MVCmuUAQCAg9olsz-qTpt5LLcfrEUIAAAD3MKtAAA=/messages/delta').get();
      const signature = response.value[0]?.body?.content || '';
      return signature;
  } catch (error) {
      console.error('Erro ao obter a assinatura do usuário:', error);
      return '';
  }
}

export async function sendMailWithAttachments(subject: string, body: string, attachmentPaths: string[], sendto: string) {
  // Ensure client isn't undefined
  if (!_userClient) {
    throw new Error('Graph has not been initialized for user auth');
  }

  // Obter a assinatura do usuário
  // const signature = await getUserSignature();
  const signature = fs.readFileSync('signature.html', 'utf-8');

  const head = "<head><style>body { font-family: 'Aptos Display', sans-serif; }</style></head>";

  // Combinar o corpo do e-mail com a assinatura
  const emailBody = `${head}${body}<br><br>${signature}`;

  // Ler os arquivos e codificá-los em base64
  const attachments = attachmentPaths.map(path => {
      const content = fs.readFileSync(path).toString('base64');
      const name = path.split('/').pop(); // Extrair o nome do arquivo do caminho
      return {
          '@odata.type': '#microsoft.graph.fileAttachment',
          name: name,
          contentBytes: content
      };
  });

  const message = {
      subject: subject,
      body: {
          contentType: 'HTML',
          content: emailBody
      },
      toRecipients: [
          {
              emailAddress: {
                  address: sendto
              }
          }
      ],
      attachments: attachments
      // ,
      // importance: 'Low', // Add email importance
      // isDeliveryReceiptRequested: true, // Request delivery receipt
      // isReadReceiptRequested: true // Request read receipt
  };

  try {
      await _userClient.api('/me/sendMail')
          .post({ message });
      console.log('Email enviado com sucesso!');
  } catch (error) {
      console.error('Erro ao enviar email:', error);
  }
}

// <MakeGraphCallSnippet>
// This function serves as a playground for testing Graph snippets
// or other code
export async function makeGraphCallAsync() {
  // INSERT YOUR CODE HERE
}
// </MakeGraphCallSnippet>
