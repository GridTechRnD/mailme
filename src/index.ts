import express from 'express';
import jwt from 'jsonwebtoken';
import multer from 'multer';
import 'dotenv/config';

import { verifyJWT } from './auth';
import { readSettings, removeAttachments, readAssets } from './files';
import initializeGraph from './graphHelper';
import initializeMail from './mail';

async function main() {
  const _graphApi = readSettings().then(initializeGraph);

  const app = express();
  const upload = multer({ dest: 'uploads/' });
  app.use(express.json());
  app.use(express.urlencoded({ extended: true }));

  const graphApi = await _graphApi;

  const mailApi = initializeMail(graphApi.userClient);

  app.post('/login', upload.none(), async (req, res) => {
    const username = process.env.USERNAME || '';
    const password = process.env.PASSWORD || '';
    const secret = process.env.SECRET || '';

    const info_req = JSON.parse(req.body.data);
    const user = info_req.user;
    const pass = info_req.password;

    if (user !== username || pass !== password) {
      return res.status(500).json({ message: 'Login inválido!' });
    }

    const id = 1;
    const token = jwt.sign({ id }, secret, {
      expiresIn: '15m'
    });

    return res.json({ auth: true, token: token });
  });

  app.post('/logout', (_req, res) => {
    return res.json({ auth: false, token: null });
  });

  app.get('/', verifyJWT, async (_req, res) => {
    return res.send({ message: 'Hello World!' });
  });

  app.post('/', verifyJWT, upload.array('file[]'), async (req, res) => {
    const { subject, body, sendto } = JSON.parse(req.body.data);

    const files = req.files as Express.Multer.File[];
    await mailApi.sendMail(subject, body, sendto, files);
    await removeAttachments(files.map((f) => f.path));

    return res.json({ message: 'Success' });
  });

  app.post('/endTicket', verifyJWT, upload.array('file[]'), async (req, res) => {
    const { sendto } = JSON.parse(req.body.data);

    const files = req.files as Express.Multer.File[];
    const body = await readAssets('endTicket.html');
    await mailApi.sendMail('SUPORTE LIVOLTEK - Final do chamado', body, sendto, files);
    await removeAttachments(files.map((f) => f.path));

    return res.json({ message: 'Success' });
  });

  app.listen(3000, () => console.log('listening on port 3000!'));
}

main().catch(console.error);
