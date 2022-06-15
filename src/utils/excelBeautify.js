/**
 * @typedef {Object} BorderDrawOptions
 * @property {{row: Number, col: Number}} start
 * @property {{row: Number, col: Number}} end
 * @property {String} borderWidth
 */
/**
 * @typedef {Object} FillCellOptions
 * @property {string} color
 * @property {string} pattern
 */

/**
 * @description Function to add a border style to a range of rows and columns
 * @param {Object} worksheet
 * @param {BorderDrawOptions} options
 * @returns {void}
 */
export const createOuterBorder = (
  worksheet = null,
  options = {
    start: { row: 1, col: 1 },
    end: { row: 1, col: 1 },
    borderWidth: "medium",
    color: "000000"
  }
) => {
  const borderStyle = {
    style: options?.borderWidth,
    color: { argb: options?.color },
  }
  for (let i = options?.start?.row; i <= options?.end?.row; i++) {
    const leftBorderCell = worksheet.getCell(i, options?.start?.col)
    const rightBorderCell = worksheet.getCell(i, options?.end?.col)
    leftBorderCell.border = {
      ...leftBorderCell.border,
      left: borderStyle,
    }
    rightBorderCell.border = {
      ...rightBorderCell.border,
      right: borderStyle,
    }
  }

  for (let i = options?.start.col; i <= options?.end.col; i++) {
    const topBorderCell = worksheet.getCell(options?.start.row, i)
    const bottomBorderCell = worksheet.getCell(options?.end.row, i)
    topBorderCell.border = {
      ...topBorderCell.border,
      top: borderStyle,
    }
    bottomBorderCell.border = {
      ...bottomBorderCell.border,
      bottom: borderStyle,
    }
  }
}

/**
 * @description
 * @param
 * @returns
 */
export const changeFontSize = (
  worksheet,
  start = { row: 1, col: 1 },
  end = { row: 1, col: 1 },
  fontSize = 11
) => {
  for (let i = start.col; i <= end.col; i++) {
    for (let j = start.row; j <= end.row; j++) {
      worksheet.getCell(j, i).font = {
        name: "Calibri",
        size: fontSize,
      }
    }
  }
}

/**
 * @description
 * @param {Object} row
 * @param {Object} column
 * @param {FillCellOptions} options
 * @returns {void}
 */
export const fillCells = (
  row = null,
  column = null,
  options = {
    color: "FFFFFF",
    pattern: "solid"
  }
) => {
  if (!row && !column) return null
  if (row)
    row.eachCell((cell, colNumber) => {
      cell.fill = {
        type: "pattern",
        pattern,
        fgColor: { argb: color },
      }
    })
  if (column)
    column.eachCell((cell, colNumber) => {
      cell.fill = {
        type: "pattern",
        pattern,
        fgColor: { argb: color },
      }
    })
}

/**
 * @description
 * @param
 * @returns
 */
export const fillCell = (cell, color) => {
  cell.fill = {
    type: "pattern",
    pattern: "darkVertical",
    fgColor: { argb: color },
  }
}

/**
 * @description
 * @param
 * @returns
 */
export const styleHeaders = (
  worksheet,
  cols = 0,
  color = "0E00CE",
  height = 40
) => {
  changeFontSize(
    worksheet,
    {
      row: 1,
      col: 1,
    },
    {
      row: 1,
      col: cols,
    },
    14
  )
  worksheet.getRow(1).height = height
  fillCells(worksheet.getRow(1), null, color, "solid")
  if (cols) {
    for (let i = 1; i <= cols; i++) {
      createOuterBorder(
        worksheet,
        {
          row: 1,
          col: i,
        },
        {
          row: 1,
          col: i,
        }
      )
    }
  }
}

const excelBeautify = {
  createOuterBorder,
  fillCells,
  fillCell,
  changeFontSize,
  styleHeaders,
}
export default excelBeautify