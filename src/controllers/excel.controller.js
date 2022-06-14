import CreateExcel from "../utils/excelGenerator"
import excelDataExtractor from "../utils/excelDataExtractor"
import excelDataProcessing from "../utils/excelDataProcessing"

const tmpData = [
  {
    "NAME":"General",
    "HEADERS_NAME":[
      "Integers",
      "Strings"
    ],
    "HEADERS_KEY":[
      "Enteros",
      "Cadenas de texto"
    ],
    "COLUMNS_WIDTH": [
      15,
      32
    ]
  }
]

const sheetFields = {
  Hoja1: [
    "Enteros",
    "Cadenas de texto"
  ],
  Hoja2: [
    "Prueba1",
    "Prueba4"
  ]
}

const getDataFromFile = async (req, res) => {
    const { filepath, sheetName, fields } = req.body
    let worksheetsData = []
    let worksheetOne = []
    //const data = await excelDataExtractor.getDataFromAllSheets(filepath, sheetFields)
    const data = await excelDataExtractor.getDataFromOneSheet(filepath, sheetName, fields)
    worksheetsData.push(excelDataProcessing.genWorksheetRows(data))
    await CreateExcel('result.xlsx', { WORKSHEETS: tmpData, DATA: worksheetsData })
    res.send(data)
}

const excelController = {
    getDataFromFile
}

export default excelController