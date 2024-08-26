// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

// <SettingsSnippet>
// const settings: AppSettings = {
//   'clientId': '',
//   'client_secret':'',
//   'directory_id':'',
//   'tenantId': '',
//   'graphUserScopes': [
//     'user.read',
//     'mail.read',
//     'mail.send'
//   ]
// };

export interface AppSettings {
  clientId: string;
  client_secret: string;
  directory_id: string;
  tenantId: string;
  graphUserScopes: string[];
}

export default settings;
// </SettingsSnippet>
