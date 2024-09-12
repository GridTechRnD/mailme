import jwt from 'jsonwebtoken';
import { Request, Response, NextFunction } from 'express';
import dotenv from 'dotenv';
dotenv.config();

interface CustomRequest extends Request {
  userId?: string;
}

export function verifyJWT(req: CustomRequest, res: Response, next: NextFunction) {
  const token = req.headers['authorization'];
  if (!token) return res.status(401).json({ auth: false, message: 'No token provided.' });

  jwt.verify(token, process.env.SECRET as jwt.Secret, (err, decoded) => {
    if (err) return res.status(500).json({ auth: false, message: 'Failed to authenticate token.' });

    req.userId = (decoded as { id: string }).id;
    next();
  });
}
