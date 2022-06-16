/**
 * @typedef {Object} BorderDrawOptions
 * @property {{row: Number, col: Number}} start
 * @property {{row: Number, col: Number}} end
 * @property {String} borderWidth
 */
/**
 * @typedef {Object} FillCell_PatternOptions
 * @property {string} pattern
 * @property {Object} fgColor
 * @property {String} fgColor.argb
 * @property {Object} bgColor
 * @property {String} bgColor.argb
 */
/**
 * @typedef {Object} GradientOptions_Stop
 * @property {string} position
 * @property {Object} color
 * @property {string} color.argb
 */
/**
 * @typedef {Object} FillCell_GradientOptions
 * @property {string} gradient
 * @property {Number} degree
 * @property {Object} center
 * @property {Number} center.left
 * @property {Number} center.top
 * @property {[GradientOptions_Stop]} stops
 */
/**
 * @typedef {Object} FontSizeOptions
 * @property {{row: Number, col: Number}} start
 * @property {{row: Number, col: Number}} end
 * @property {Number} fontSize
 * @property {string} fontType
 */

/**
 * @description Function to add a border style to a range of rows and columns
 * @param {Object} worksheet
 * @param {BorderDrawOptions} options
 * @returns {boolean}
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
 * @param {Object} worksheet
 * @param {FontSizeOptions} options
 * @returns {boolean}
 */
export const changeFontProperties = (
  worksheet,
  options = {
    start: { row: 1, col: 1 },
    end: { row: 1, col: 1 },
    fontSize: 12,
    fontType: "Calibri"
  }
) => {
  for (let i = options?.start.col; i <= options?.end.col; i++) {
    for (let j = options?.start.row; j <= options?.end.row; j++) {
      worksheet.getCell(j, i).font = {
        name: options?.fontType,
        size: options?.fontSize,
      }
    }
  }
}

/**
 * @description
 * @param {Object} row
 * @param {Object} column
 * @param {FillCellOptions} options
 * @returns {boolean}
 */
export const fillCellsWithPattern = (
  row = null,
  column = null,
  options = {
    pattern: "solid",
    fgColor: { argb: "000000" },
    bgColor: { argb: "ffffff" },
  }
) => {
  if (!row && !column) return null
  if (row)
    row.eachCell((cell, colNumber) => {
      console.log(cell)
      cell.fill = { type:"pattern", ...options }
      console.log(cell)
    })
  if (column)
    column.eachCell((cell, colNumber) => {
      cell.fill = { type:"pattern", ...options }
    })
}

/**
 * @description
 * @param {Object} cell
 * @param {FillCell_PatternOptions} options
 * @returns {boolean}
 */
export const fillCellWithPattern = (cell,
  options = {
    pattern: "solid",
    fgColor: { argb: "000000" },
    bgColor: { argb: "ffffff" },
  }
) => {
  if (!cell) return null
  cell.fill = { type:"pattern", ...options }
}

/**
 * @description
 * @param
 * @returns
 */
export const styleHeaders = (
  worksheet,
  options = {
    columnsNo: 0,
    fgColor: "000000",
    bgColor: "ffffff",
    rowHeight: 40

  }
) => {
  changeFontProperties(
    worksheet,
    {
      start: {
        row: 1,
        col: 1,
      },
      end: {
        row: 1,
        col: options.columnsNo,
      },
      fontSize: 14,
      fontType: "Calibri",
    }
  )
  worksheet.getRow(1).height = options.rowHeight
  fillCellsWithPattern(worksheet.getRow(1), null,
    {
      pattern: "solid",
      fgColor: { argb: options.fgColor },
      bgColor: { argb: options.bgColor },
    }
  )
  if (options.columnsNo) {
    for (let i = 1; i <= options.columnsNo; i++) {
      createOuterBorder(
        worksheet,
        {
          start: {
            row: 1,
            col: i,
          },
          end: {
            row: 1,
            col: i,
          },
          borderWidth: "medium",
        }
      )
    }
  }
}

const excelBeautify = {
  createOuterBorder,
  fillCellsWithPattern,
  fillCellWithPattern,
  changeFontProperties,
  styleHeaders,
}
export default excelBeautify