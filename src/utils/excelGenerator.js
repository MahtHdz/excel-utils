import ExcelJS from "exceljs"

import worksheetDefaultOptions from '../config/worksheetConfig.json'

/**
 * @typedef {Object} worksheet
 * @property {String} NAME - name of the worksheet
 * @property {[String]} HEADERS_NAME - name of the headers
 * @property {[String]} HEADERS_KEY - key of the headers
 * @property {[String]} COLUMNS_WIDTH - width of the columns
 */

/**
 * @description Function to create an excel file report
 * @param {String} filepath - path of the file to be exported
 * @param {Object} excelConstructor - object with the excel constructor parameters
 * @param {[worksheet]} excelConstructor.WORKSHEETS_CONFIG - worksheets structure configuration
 * @param {[Object]} excelConstructor.WORKSHEETS_DATA - worksheets data
 * @param {[Object]} excelConstructor.WORKSHEETS_OPTIONS - worksheets constructor configuration
 * @param {Function} customStyleFunction - custom function to style the excel sheets
 * @returns {Promise<boolean>} Return true if the file was created successfully
 */
const CreateExcel = async (filepath, excelConstructor, customStyleFunction) => {
  let worksheets = []
  let worksheetOptions = null
  let worksheetsOptionsFlag = false

  // Initialize workbook
  const workbook = new ExcelJS.Workbook()

  // Verify if worksheets options are defined
  if (excelConstructor.WORKSHEETS_OPTIONS) worksheetsOptionsFlag = true

  // Building the worksheets
  excelConstructor.WORKSHEETS_CONFIG.forEach((worksheet, index) => {
    worksheetOptions = worksheetsOptionsFlag ? excelConstructor?.WORKSHEETS_OPTIONS[index] : worksheetDefaultOptions.GENERAL
    worksheets.push(workbook.addWorksheet(worksheet.NAME, worksheetOptions))
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
  excelConstructor.WORKSHEETS_DATA.forEach((worksheet, index) => {
    worksheet.forEach(row => {
      let excelRow = worksheets[index].addRow({
        ...row
      })
    })
  })

  /********* Styling the worksheets *********/
  if(customStyleFunction) customStyleFunction(worksheets)

  await workbook.xlsx.writeFile(filepath)
  console.log("File is written")
  return true
}

export default CreateExcel