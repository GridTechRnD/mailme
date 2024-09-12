import 'isomorphic-fetch';
import { DeviceCodeCredential, DeviceCodePromptCallback, UsernamePasswordCredential } from '@azure/identity';
import { Client } from '@microsoft/microsoft-graph-client';
import { User } from '@microsoft/microsoft-graph-types';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';

import type AppSettings from './appSettings';

export default async function initializeGraph(settings: AppSettings, deviceCodePrompt?: DeviceCodePromptCallback) {
  const username = process.env.USERNAME;
  const password = process.env.PASSWORD_OUTLOOK;

  if (!username || !password) {
    throw new Error('Username and password must be provided');
  }

  const credential = new UsernamePasswordCredential(settings.directory_id, settings.clientId, username, password);

  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: settings.graphUserScopes
  });

  const userClient = Client.initWithMiddleware({
    authProvider: authProvider
  });

  const deviceCodeCredential = new DeviceCodeCredential({
    clientId: settings.clientId,
    tenantId: settings.tenantId,
    userPromptCallback: deviceCodePrompt
  });

  const getUserToken = async (): Promise<string> => {
    const response = await deviceCodeCredential.getToken(settings.graphUserScopes);
    return response.token;
  };

  const getUser = async (): Promise<User> => {
    return userClient.api('/me').select(['displayName', 'mail', 'userPrincipalName']).get();
  };

  return {
    userClient,
    getUser,
    getUserToken
  };
}
