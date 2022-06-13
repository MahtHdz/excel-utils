import ExcelJS from "exceljs"

import StyleExcel from "./Style"
import WORKSHEETS from "./config/worksheets.json"
import worksheetConfig from '../config/worksheetConfig.json'

/**
 * @description Function to create an excel file report
 * @param {Array} dataArr Array of data to be exported
 * @param {String} filename Name of the file to be exported
 * @param {String} filepath Path of the file to be exported
 * @param {String} userId User Id of the user who is exporting the file
 * @returns {Promise<boolean>} Return true if the file was created successfully
 */
const CreateExcel = async (dataArr, filename, filepath, userId, excelConstructor) => {
  // Initialize workbook
  const workbook = new ExcelJS.Workbook()

  // Creation of worksheets
  let worksheets = []
  WORKSHEETS.WORKSHEETS.forEach((worksheet, index) => {
    worksheets.push(workbook.addWorksheet(worksheet.NAME, worksheetConfig.GENERAL))
    // Setting headers
    let columns = []
    worksheet.CELLS.PLD_CELL_NAMES.map((name, i) => {
      columns.push({
        header: name,
        key: worksheet.CELLS.PLD_CELL_KEYS[i],
        width: worksheet.CELLS.PLD_CELL_WIDTH[i]
      })
    })
    worksheets[index].columns=columns
  })

  // Initializing the operations accumulation value index
  let opsAccumulationValueIndex = 0
  dataArr[0].forEach(op => {
    /*
      Check if op const is an array.
      If it is an array, it means that
      the op const is a group of operations
    */
    if(op.constructor === Array) {
      let row = null
      const worksheetIndexs = [0, 3]
      /*
        Adding operations group to worksheet and styling it.

        Note: The prevRowCount and currentRowCount variables
        are to control the styling and merge of the operations
        group. Consider that the prevRowCount need and extra
        position because is conting the first operation in the
        operation group.
      */
      worksheetIndexs.forEach(index => {
        let prevRowCount = worksheets[index].actualRowCount + 1
        op.forEach(operation => {
          row = worksheets[index].addRow({
              rfc: operation.RFC,
              razonSocial: operation.Razon_Social,
              referencia: operation.Referencia,
              fechaMovimiento: operation.Fecha_Movimiento,
              montoPagado: operation.Monto_Pagado,
          })
          // Styling the current row
          StyleExcel.fillCells(row, null,'D8CA00', 'darkVertical')
        })
        // Getting the current row count
        let currentRowCount = worksheets[index].actualRowCount
        // Merging the operations group
        worksheets[index].mergeCells(prevRowCount, 6, currentRowCount, 6)
        // Getting of the operations accumulation value with the index in the dataArr
        row.getCell(6).value = dataArr[1][opsAccumulationValueIndex]
        // Styling the operations accumulation value
        StyleExcel.fillCell(row.getCell(6), 'EB5353') // Alternative color = 'D30000'
        StyleExcel.createOuterBorder(worksheets[index], {row: prevRowCount, col: 1}, {row: currentRowCount, col: 6})
      })
      opsAccumulationValueIndex++
    } else {
      /*
        Note: The operations written in this conditional
        are the ones that exceed the warning threshold.
      */
      const worksheetIndexs = [0, 2]
      worksheetIndexs.forEach(index => {
        let row = worksheets[index].addRow({
          rfc: op.RFC,
          razonSocial: op.Razon_Social,
          referencia: op.Referencia,
          fechaMovimiento: op.Fecha_Movimiento,
          montoPagado: op.Monto_Pagado,
        })
        let currentRowCount = worksheets[index].actualRowCount
        StyleExcel.fillCells(row, null,'EB5353', 'darkVertical') // Alternative color = 'D30000'
        StyleExcel.createOuterBorder(worksheets[index], {row: currentRowCount, col: 1}, {row: currentRowCount, col: 5})
      })
    }
  })
  /**
   * Writing the operations tha exceed the identification threshold
   * and the paid amount is less than the warning threshold.
   */
  dataArr[2].forEach(op => {
    let row = worksheets[1].addRow({
      rfc: op.RFC,
      razonSocial: op.Razon_Social,
      referencia: op.Referencia,
      fechaMovimiento: op.Fecha_Movimiento,
      montoPagado: op.Monto_Pagado,
    })
    // Styling the current row
    let currentRowCount = worksheets[1].actualRowCount
    StyleExcel.fillCells(row, null,'D8CA00', 'darkVertical')
    StyleExcel.createOuterBorder(worksheets[1], {row: currentRowCount, col: 1}, {row: currentRowCount, col: 5})
  })

  /********* Styling the worksheets *********/
  worksheets.forEach((worksheet, index) => {
    // Setting headers style
    StyleExcel.styleHeaders(worksheet, worksheet.columns.length)
    // Center all text
    WORKSHEETS.WORKSHEETS[index].CELLS.PLD_CELL_KEYS.forEach(key => {
      const column = worksheet.getColumnKey(key)
      if(key !== 'razonSocial') column.eachCell((cell, rowNumber) => {
        cell.alignment = {
          vertical: 'middle',
          horizontal: 'center',
          wrapText: true
        }
      })
      else column.eachCell((cell, rowNumber) => {
        if(rowNumber===1) cell.alignment = {
          vertical: 'middle',
          horizontal: 'center',
          wrapText: true
        }
        return
      })
    })
  })
  // Write the xlsx file
  const obj = {
    name: filename,
    path: filepath,
    userId
  }
  //await axios.post('http://localhost:4000/api/clientReport/saveOne', obj)
  await workbook.xlsx.writeFile(filepath)
  console.log("File is written")
  return true
}

export default CreateExcel