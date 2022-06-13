import express from 'express'

import pkg from '../package.json'

import excelRouter from './routes/excel.routes'


const app = express()

app.use(express.json())

app.set('pkg', pkg)
app.get('/', (req, res) => {
    res.json({
        author      : app.get('pkg').author,
        description : app.get('pkg').description,
        version     : app.get('pkg').version
    })
})

app.use('/api/excel', excelRouter)

export default app