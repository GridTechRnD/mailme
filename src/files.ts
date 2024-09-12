import fs from 'node:fs/promises';
import type AppSettings from './appSettings';

export async function readSettings(): Promise<AppSettings> {
  const file = await fs.readFile('appSettings.json', 'utf-8');
  return JSON.parse(file) as AppSettings;
}

export interface Attachment {
  '@odata.type': '#microsoft.graph.fileAttachment';
  name: string;
  contentBytes: string;
}

export async function readAttachments(attachments: Express.Multer.File[]): Promise<Attachment[]> {
  return await Promise.all(
    attachments.map(async (attc) => {
      const path = attc.path;
      const name = attc.originalname;

      const content = await fs.readFile(path, 'base64');

      return {
        '@odata.type': '#microsoft.graph.fileAttachment',
        name: name,
        contentBytes: content
      };
    })
  );
}

export async function readAssets(filename: string): Promise<string> {
  return await fs.readFile(`assets/${filename}`, 'utf-8');
}

export async function removeAttachments(attachments: string[]) {
  await Promise.all(
    attachments.map(async (path) => {
      fs.rm(path);
    })
  );
}
