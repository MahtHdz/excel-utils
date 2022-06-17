import excelBeautify from "./excelBeautify"

/**
 * @description
 * @param {[Object]} worksheets
 * @returns {boolean}
 */
const myCustomStyleFunction = (worksheets) => {
  let headersConfig = {
    fillColor: "140484",
    fontColor: "ffffff",
    borderColor: "ffffff",
    rowHeight: 40
  }
  worksheets.forEach((worksheet, index) => {
    // Setting headers style
    headersConfig.columnsNo = worksheet.columns.length
    excelBeautify.styleHeaders(worksheet, headersConfig)
    // Center all text
    excelBeautify.cellAlignmentForAllColumns(worksheet, {
      horizontal: "center",
      vertical: "middle",
      wrapText: true
    })
  })

}

export default myCustomStyleFunction