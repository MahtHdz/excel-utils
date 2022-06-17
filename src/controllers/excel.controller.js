import CreateExcel from "../utils/excelGenerator"
import myCustomStyleFunction from "../utils/customStyle"
import excelDataExtractor from "../utils/excelDataExtractor"
import excelDataProcessing from "../utils/excelDataProcessing"

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
  },
  {
    "NAME":"Test",
    "HEADERS_NAME":[
      "Prueba 342",
      "SET de prueba"
    ],
    "HEADERS_KEY":[
      "Prueba1",
      "Prueba4"
    ],
    "COLUMNS_WIDTH": [
      15,
      15
    ]
  }
]

const worksheetsConfig = [
  {
    "pageSetup": {
      "paperSize": 9,
      "orientation": "landscape"
    }
  },
  {
    "pageSetup": {
      "paperSize": 9,
      "orientation": "landscape"
    }
  }
]

const getDataFromFile = async (req, res) => {
    const { filepath, sheetName, fields } = req.body
    let worksheetsData = []
    let worksheetOne = []
    const data = await excelDataExtractor.getDataFromAllSheets(filepath, sheetFields)
    data.forEach(sheet => {
      worksheetsData.push(excelDataProcessing.genSheetRows(sheet.fieldsData))
    })
    await CreateExcel('result.xlsx', {
      WORKSHEETS_CONFIG: tmpData,
      WORKSHEETS_DATA: worksheetsData,
      WORKSHEETS_OPTIONS: worksheetsConfig
    }, myCustomStyleFunction)
    //const data = await excelDataExtractor.getDataFromOneSheet(filepath, sheetName, fields)
    //worksheetsData.push(excelDataProcessing.genWorksheetRows(data))
    //await CreateExcel('result.xlsx', { WORKSHEETS: tmpData, DATA: worksheetsData })
    res.send(data)
}

const excelController = {
    getDataFromFile
}

export default excelController