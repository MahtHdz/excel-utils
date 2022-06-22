import { Router } from 'express'

import excelController from '../controllers/excel.controller'

const excelRouter = Router()

excelRouter.post('/', excelController.getDataFromFile)
excelRouter.post('/test', excelController.testArrays)

export default excelRouter