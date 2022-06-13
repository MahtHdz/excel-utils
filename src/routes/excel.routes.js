import { Router } from 'express'

import excelController from '../controllers/excel.controller'

const excelRouter = Router()

excelRouter.post('/', excelController.getDataFromFile)

export default excelRouter