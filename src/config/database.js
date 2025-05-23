import dotenv from 'dotenv';
dotenv.config();

const timeout = parseInt(process.env.DB_TIMEOUT, 10) || 90000;
const encrypt = process.env.DB_ENCRYPT === 'true';

export const config = {
    server: process.env.DB_SERVER,
    authentication: {
        type: 'default',
        options: {
            userName: process.env.DB_USERNAME,
            password: process.env.DB_PASSWORD
        }
    },
    options: {
        connectionTimeout: timeout,
        requestTimeout: timeout,
        encrypt: encrypt,
        database: process.env.DB_NAME
    }
};
