import ExcelJS from "exceljs"

import worksheetConfig from '../config/worksheetConfig.json'

/**
 * @typedef {Object} worksheet
 * @property {String} NAME - name of the worksheet
 * @property {[String]} HEADERS_NAME - name of the headers
 * @property {[String]} HEADERS_KEY - key of the headers
 * @property {[String]} COLUMNS_WIDTH - width of the columns
 */

/**
 * @description Function to create an excel file report
 * @param {String} filepath Path of the file to be exported
 * @param {Object} excelConstructor
 * @param {[worksheet]} excelConstructor.WORKSHEETS
 * @param {[Object]} excelConstructor.DATA
 * @returns {Promise<boolean>} Return true if the file was created successfully
 */
const CreateExcel = async (filepath, excelConstructor) => {
  // Initialize workbook
  const workbook = new ExcelJS.Workbook()

  // Building the worksheets
  let worksheets = []
  excelConstructor.WORKSHEETS.forEach((worksheet, index) => {
    worksheets.push(workbook.addWorksheet(worksheet.NAME, worksheetConfig.GENERAL))
    // Setting headers
    let columns = []
    worksheet.HEADERS_NAME.map((name, i) => {
      columns.push({
        header: name,
        key: worksheet.HEADERS_KEY[i],
        width: worksheet.COLUMNS_WIDTH[i]
      })
    })
    worksheets[index].columns=columns
  })

  // Data injection in the worksheets
  excelConstructor.DATA.forEach((worksheet, index) => {
    worksheet.forEach(row => {
      let excelRow = worksheets[index].addRow({
        ...row
      })
    })
  })

  await workbook.xlsx.writeFile(filepath)
  console.log("File is written")
  return true
}

export default CreateExcel