import ExcelJS from "exceljs"
/**
* @description This function is used to get the data from selected fields in a excel file
* @param {String} filepath - path of the excel file.
* @param {String} sheetName - sheet name where the data is located.
* @param {String} fields - fields selected to get the file columns.
* @returns {Promise<Array>} - returns a promise with the data in the excel file.
*/
const getData = async (filepath, sheetName, fields) => {
  let clientData = {}
  let fieldsArray = []
  let workSheet = null
  let fieldsFlag = []
  const workbook = new ExcelJS.Workbook()
  // Getting the layout headers
  const fieldsName = Object.values(fields)
  // Initializing the fields flag array
  fieldsName.map((field, index) => (fieldsFlag[index] = false))
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
      if (fieldsFlag[count] && cell.value !== null) clientData[currentField].push(cell.value)
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

const excelDataExtractor = {
    getData
}

export default excelDataExtractor