import { Router } from 'express'
import multer from 'multer'
import PublicFileController from './app/controllers/PublicFileController'
import GenerateExcelController from './app/controllers/GenerateExcelController'

const routes = new Router()

const upload = multer({ dest: 'public/uploads/' })

routes.post('/generate/excel', GenerateExcelController.store)
routes.get('/get-file/:name', PublicFileController.show)

export default routes