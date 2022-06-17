import ExcelJS from "exceljs"

/**
* @description This function is used to get the data from selected fields in a excel file
* @param {String} filepath - path of the excel file.
* @param {String} sheetName - sheet name where the data is located.
* @param {String} fields - fields selected to get the file columns.
* @returns {Promise<Array>} - returns a promise with the data in the excel file.
*/
const getDataFromOneSheet = async (filepath, sheetName, fields) => {
  let clientData = {}
  let fieldsArray = []
  let workSheet = null
  let fieldsFlag = []
  const workbook = new ExcelJS.Workbook()
  // Getting the layout headers
  const fieldsName = Object.values(fields)
  // Initializing the fields flag array
  fieldsName.forEach((field, index) => (fieldsFlag[index] = false))
  // Trying to get the workbook from excel file
  try { await workbook.xlsx.readFile(filepath) }
  catch (error) { return console.error(error) }
  // Getting worksheet from the workbook
  workSheet = workbook.getWorksheet(sheetName)
  // Getting the raw data from fields and saving them into an array
  for (let i = 1; i <= workSheet.actualColumnCount; i++) { fieldsArray.push(workSheet.getColumn(i)) }
  /*
    Filtering and sorting the data in fieldsArray
    and saving it into an object
  */
  // Counter to verify if all the fields are found
  let count = 0
  fieldsArray.forEach((column) => {
    let currentField = ""
    // If count is less than the length of the fieldsFlags, then there's still data to be saved
    if (count === fieldsFlag.length) return
    // Getting the data from the current column
    column.eachCell((cell, rowNumber) => {
      /*
        If the field title was found and still there
        are data to be saved in the current column,
        save the data from current cell, if not,
        find the field title in the current cell.
      */
      if (fieldsFlag[count] && cell.value !== null) {
        if (cell.value instanceof Object)
          clientData[currentField].push(cell.value.result)
        else
          clientData[currentField].push(cell.value)
      }
      else fieldsName.forEach((name) => {
        // Finding the field title in the current cell
        if (cell.value === name) {
          fieldsFlag[count] = true
          clientData[name] = []
          currentField = name
          return
        }
      })
    })
    if(fieldsFlag[count]) count++
  })
  return clientData
}

/**
* @description This function is used to get the data from selected fields in a excel file
* @param {String} filepath - path of the excel file.
* @param {String} sheetsFields - sheet name where the data is located.
* @returns {Promise<Array>} - returns a promise with the data in the excel file.
*/
const getDataFromAllSheets = async (filepath, sheetsFields) => {
  let sheetsFieldsData = []
  //let worksheets = []
  let sheetsFieldsFlag = []
  let sheetsData = []
  const workbook = new ExcelJS.Workbook()
  // Getting the layout headers
  const sheetsFieldsName = Object.entries(sheetsFields)
  // Initializing the fields flag array
  sheetsFieldsName.forEach(sheet => {
    let fieldsFlag = []
    sheet[1].forEach(() => { fieldsFlag.push(false) })
    sheetsFieldsFlag.push(fieldsFlag)
  })
  // Trying to get the workbook from excel file
  try { await workbook.xlsx.readFile(filepath) }
  catch (error) { return console.error(error) }
  // Getting worksheet from the workbook
  workbook.eachSheet((sheet, sheetId) => {
    let fieldsArray = []
    // Getting the raw data from fields and saving them into an array
    for (let i = 1; i <= sheet.actualColumnCount; i++) {
      fieldsArray.push(sheet.getColumn(i))
    }
    sheetsData.push([sheet.name, fieldsArray])
    //worksheets.push(sheet)
  })
  sheetsData.forEach(sheet => {
    let fieldsData = {}
    sheetsFieldsName.forEach((sheetFields, index) => {
      if (sheet[0] === sheetFields[0]) {
        // Counter to verify if all the fields are found
        let count = 0
        sheet[1].forEach((column) => {
          let currentField = ""
          // If count is less than the length of the fieldsFlags, then there's still data to be saved
          if (count === sheetsFieldsFlag[index].length) return
          // Getting the data from the current column
          column.eachCell((cell, rowNumber) => {
            /*
              If the field title was found and still there
              are data to be saved in the current column,
              save the data from current cell, if not,
              find the field title in the current cell.
            */
            if (sheetsFieldsFlag[index][count] && cell.value !== null) {
              if (cell.value instanceof Object)
                fieldsData[currentField].push(cell.value.result)
              else
                fieldsData[currentField].push(cell.value)
            }
            else sheetFields[1].forEach((name) => {
              // Finding the field title in the current cell
              if (cell.value === name) {
                sheetsFieldsFlag[index][count] = true
                fieldsData[name] = []
                currentField = name
                return
              }
            })
          })
          if(sheetsFieldsFlag[index][count]) count++
        })
      }
    })
    sheetsFieldsData.push({ sheetName:sheet[0], fieldsData })
  })
  return sheetsFieldsData
}

const excelDataExtractor = {
    getDataFromOneSheet,
    getDataFromAllSheets
}

export default excelDataExtractor