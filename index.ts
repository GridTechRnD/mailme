// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

// <ProgramSnippet>
import { DeviceCodeInfo } from '@azure/identity';
import { Message } from '@microsoft/microsoft-graph-types';
import express from 'express';
import settings, { AppSettings } from './appSettings';
import * as graphHelper from './graphHelper';
import uploads from './uploads';
import jwt from 'jsonwebtoken';
import dotenv from 'dotenv';
import { appendFile, unlink } from 'fs';
import multer from 'multer';
dotenv.config();


async function verifyToken(req: express.Request, res: express.Response) {
  const token = req.headers['authorization'];
  if (!token) return res.status(401).json({ auth: false, message: 'No token provided.' });

  await jwt.verify(token, process.env.SECRET as jwt.Secret, function(err: any, decoded: any) {
    if (err) return res.status(500).json({ auth: false, message: 'Failed to authenticate token.' });
  });
}


async function main() {

  // Initialize Graph
  initializeGraph(settings);
  
  const app = express();
  const upload = multer();

  // app.use(express.urlencoded({ extended: true }));

  // Middleware para processar JSON e URL-encoded data
  app.use(express.json());
  app.use(express.urlencoded({ extended: true }));

  //User Authentication
  app.post('/login', upload.none(), async (req, res, next) => {
    const username : string = process.env.USERNAME || '';
    const password : string = process.env.PASSWORD || '';
    const secret : string = process.env.SECRET || '';
    
    // Ajuste para ler o body corretamente
    const info_req = JSON.parse(req.body.data);
    const user = info_req.user;
    const pass = info_req.password

    if(user === username && pass === password){
      //auth ok
      const id = 1; //esse id viria do banco de dados
      const token = await jwt.sign({ id }, secret, {
        expiresIn: 900000 // expires in 15min
      });
      return res.json({ auth: true, token: token });
    }
    
    res.status(500).json({message: 'Login inválido!'});
  })

  app.post('/logout', function(req, res) {
    res.json({ auth: false, token: null });
  })

  app.get('/',async (req,res) => {
      await verifyToken(req, res);
      res.send('Hello World');
  })

  app.post('/', uploads.array('file[]'), async (req, res)  => {
      try {
        await verifyToken(req, res);
        let uploaded_files: string[] = [];
        if (Array.isArray(req.files)) {
          for (const file of req.files) {
            uploaded_files.push(file.path);
          }
        }
        const { subject, body, sendto } = JSON.parse(req.body.data); // Analisa o campo data como JSON
        await sendMailAsync(subject, body, sendto, uploaded_files ? uploaded_files : []);
        res.send(`Arquivo(s) enviado(s) com sucesso, Subject: ${subject}, Body: ${body}`);
        // Remove files from the /opt/graphtutorial/uploads/ directory
        for (const file of uploaded_files) {
          unlink(file, (err) => {
            if (err) {
              console.log(`Error removing file: ${err}`);
            } else {
              console.log(`File removed: ${file}`);
            }
          });
        }
      } catch (error: any) {
          res.status(500).send(`Erro ao processar a solicitação: ${error}`);
      }
    })

  app.listen(3000, () => console.log('listening on port 3000!'));

  // await sendMailAsync('Teste', 'Teste', 'davi.santos@hexing.com.br', []);
}

// <InitializeGraphSnippet>
function initializeGraph(settings: AppSettings) {
  graphHelper.initializeGraphForUserAuth(settings, (info: DeviceCodeInfo) => {
    // Display the device code message to
    // the user. This tells them
    // where to go to sign in and provides the
    // code to use.
    console.log(info.message);
  });
}
// </InitializeGraphSnippet>

// <GreetUserSnippet>
async function greetUserAsync() {
  try {
    const user = await graphHelper.getUserAsync();
    console.log(`Hello, ${user?.displayName}!`);
    // For Work/school accounts, email is in mail property
    // Personal accounts, email is in userPrincipalName
    console.log(`Email: ${user?.mail ?? user?.userPrincipalName ?? ''}`);
  } catch (err) {
    console.log(`Error getting user: ${err}`);
  }
}
// </GreetUserSnippet>

// <DisplayAccessTokenSnippet>
async function displayAccessTokenAsync() {
  try {
    const userToken = await graphHelper.getUserTokenAsync();
    console.log(`User token: ${userToken}`);
  } catch (err) {
    console.log(`Error getting user access token: ${err}`);
  }
}
// </DisplayAccessTokenSnippet>

// <ListInboxSnippet>
async function listInboxAsync() {
  try {
    const messagePage = await graphHelper.getInboxAsync();
    const messages: Message[] = messagePage.value;

    // Output each message's details
    for (const message of messages) {
      console.log(`Message: ${message.subject ?? 'NO SUBJECT'}`);
      console.log(`  From: ${message.from?.emailAddress?.name ?? 'UNKNOWN'}`);
      console.log(`  Status: ${message.isRead ? 'Read' : 'Unread'}`);
      console.log(`  Received: ${message.receivedDateTime}`);
    }

    // If @odata.nextLink is not undefined, there are more messages
    // available on the server
    const moreAvailable = messagePage['@odata.nextLink'] != undefined;
    console.log(`\nMore messages available? ${moreAvailable}`);
  } catch (err) {
    console.log(`Error getting user's inbox: ${err}`);
  }
}
// </ListInboxSnippet>

const isValidEmail = (email: string): boolean => {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
};

// <SendMailSnippet>
async function sendMailAsync(subject: string, body: string, sendto : string, attachments: string[] = []) {
  try {
    // Send mail to the signed-in user
    // Get the user for their email address
    const user = await graphHelper.getUserAsync();
    // const userEmail = user?.mail ?? user?.userPrincipalName;

    if (!sendto) {
      console.log('Couldn\'t get your email address, canceling...');
      return;
    }else if (!isValidEmail(sendto)) {
      console.log('Recipient is not a valid email address, canceling...');
    }

    if (attachments.length > 0) {
      await graphHelper.sendMailWithAttachments(subject,body,attachments,sendto);
    }else{
      await graphHelper.sendMailAsync(subject,body, sendto);
    }
  } catch (err) {
    console.log(`Error sending mail: ${err}`);
  }
}
// </SendMailSnippet>

// <MakeGraphCallSnippet>
async function makeGraphCallAsync() {
  try {
    await graphHelper.makeGraphCallAsync();
  } catch (err) {
    console.log(`Error making Graph call: ${err}`);
  }
}

main();

// </MakeGraphCallSnippet>
