import app from './app.js'
import dotenv from 'dotenv';
dotenv.config();

app.listen(process.env.APP_PORT || 3333, () => {
    console.log(`✅ Server running on port ${process.env.APP_PORT || 3333}`)
})
