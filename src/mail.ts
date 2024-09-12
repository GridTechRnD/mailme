import { Client, PageCollection } from '@microsoft/microsoft-graph-client';

import * as files from './files';

export default function initializeMail(userClient: Client) {
  const getInbox = async (): Promise<PageCollection> => {
    return userClient
      .api('/me/mailFolders/inbox/messages')
      .select(['from', 'isRead', 'receivedDateTime', 'subject'])
      .top(25)
      .orderby('receivedDateTime DESC')
      .get();
  };

  const sendMail = async (subject: string, body: string, sendTo: string, attachments: Express.Multer.File[]) => {
    const atts = files.readAttachments(attachments);
    //const head = files.readAssets('head.html');
    const signature = files.readAssets('signature.html');

    const message = await Promise.all([atts, signature]).then(([atts, signature]) => {
      const emailBody = `${body}<br><br>${signature}`;

      return {
        subject: subject,
        body: {
          contentType: 'html',
          content: emailBody
        },
        toRecipients: [
          {
            emailAddress: {
              address: sendTo
            }
          }
        ],
        attachments: atts
      };
    });

    return userClient.api('me/sendMail').post({
      message: message
    });
  };

  const isValidEmail = (email: string): boolean => {
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(email);
  };

  return {
    getInbox,
    isValidEmail,
    sendMail
  };
}
