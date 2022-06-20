import CreateExcel from "../utils/excelGenerator"
import myCustomStyleFunction from "../utils/customStyle"
import excelDataExtractor from "../utils/excelDataExtractor"
import excelDataProcessing from "../utils/excelDataProcessing"

const getDataFromFile = async (req, res) => {
    const { filepath } = req.body
    let WORKSHEETS_CONFIG = []
    let WORKSHEETS_DATA = []
    const data = await excelDataExtractor.getAllDataFromFile(filepath)
    data.forEach(sheet => {
      let sheetConfig = {}
      const content = excelDataProcessing.genSheetRowsWithContext(sheet.fieldsData)
      WORKSHEETS_DATA.push(content.rows)
      sheetConfig.NAME = sheet.sheetName
      sheetConfig.HEADERS_NAME = content.fields
      sheetConfig.HEADERS_KEY = content.fields
      sheetConfig.COLUMNS_WIDTH = Array.from({ length: content.fields.length }, () => 15)
      WORKSHEETS_CONFIG.push(sheetConfig)
    })
    await CreateExcel('result.xlsx', {
      WORKSHEETS_CONFIG,
      WORKSHEETS_DATA,
    }, myCustomStyleFunction)
    res.send("Done")
}

const excelController = {
    getDataFromFile
}

export default excelController