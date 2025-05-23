import express from 'express'
import cors from 'cors'
import routes from './routes.js'

class App {
    constructor() {
        this.server = express()

        this.middlewares()
        this.routes()
    }

    middlewares() {
        this.server.use(cors())
        this.server.use(
            cors({
                origin: '*',
                methods: 'GET,HEAD,PUT,PATCH,POST,DELETE',
                credentials: true, // This option is important for handling cookies and authentication headers
            })
        )
    }

    routes() {
        this.server.use(routes)
    }
}

export default new App().server
