/**
 * @typedef {Object} cellAlignmentOptions
 * @property {String} vertical - Vertical alignment
 * @property {String} horizontal - Horizontal alignment
 * @property {boolean} wrapText - Wrap text
 */
/**
 * @typedef {Object} BorderDrawOptions
 * @property {{row: Number, col: Number}} start
 * @property {{row: Number, col: Number}} end
 * @property {String} borderWidth
 * @property {Object} color
 * @property {String} color.argb
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
 * @typedef {Object} FontOptions
 * @property {{row: Number, col: Number}} start
 * @property {{row: Number, col: Number}} end
 * @property {string} fontType
 * @property {Number} fontFamily
 * @property {string} scheme
 * @property {Number} charset
 * @property {Number} fontSize
 * @property {Object} color
 * @property {string} color.argb
 * @property {boolean} bold
 * @property {boolean} italic
 * @property {boolean | string} underline
 * @property {boolean} strike
 * @property {boolean} outline
 * @property {string} vertAlign
 */
/**
 * @typedef {Object} StyleHeadersOptions
 * @property {Object} columnsNo
 * @property {Object} fillColor
 * @property {Object} borderColor
 * @property {Object} fontType
 * @property {Object} fontColor
 * @property {Object} fontFamily
 * @property {Object} fontSize
 * @property {Object} rowHeight
 */

/**
 * @description
 * @param {Object} worksheet
 * @param {cellAlignmentOptions} options
 * @returns {boolean}
 */
export const cellAlignmentForAllColumns = (
  worksheet,
  options = {
    vertical: 'middle',
    horizontal: 'center',
    wrapText: true
}) => {
  for (let i = 1; i <= worksheet.actualColumnCount; i++) {
    worksheet.getColumn(i).eachCell((cell, rowNumber) => {
      cell.alignment = options
    })
  }
}

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
    color: { argb:"000000" }
  }
) => {
  const borderStyle = {
    style: options?.borderWidth,
    color: options?.color,
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
 * @param {FontOptions} options
 * @returns {boolean}
 */
export const changeFontProperties = (
  worksheet,
  options = {
    start: { row: 1, col: 1 },
    end: { row: 1, col: 1 },
    fontSize: 12,
    fontType: "Calibri",
    family: fontFamily ?? 1,
    color: { argb: "000000" }
  }
) => {
  const { start, end, fontSize, fontType, fontFamily, ...rest } = options
  for (let i = start?.col; i <= end?.col; i++) {
    for (let j = start?.row; j <= end?.row; j++) {
      worksheet.getCell(j, i).font = {
        name: fontType ?? "Calibri",
        size: fontSize ?? 12,
        family: fontFamily ?? 1,
        ...rest,
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
      cell.fill = { type:"pattern", ...options }
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
 * @param {Object} worksheet
 * @param {StyleHeadersOptions} options
 * @returns {boolean}
 */
export const styleHeaders = (
  worksheet,
  options = {
    columnsNo: 0,
    fillColor: "ffffff",
    borderColor: "000000",
    fontType: "Calibri",
    fontColor: "000000",
    fontFamily: 1,
    fontSize: 14,
    rowHeight: 20

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
        col: options?.columnsNo ?? 0,
      },
      fontSize: options?.fontSize ?? 14,
      fontType: options?.fontType ?? "Calibri",
      fontFamily: options?.fontFamily ?? 1,
      color: { argb: options?.fontColor ?? "000000" },
    }
  )
  worksheet.getRow(1).height = options?.rowHeight ?? 20
  fillCellsWithPattern(worksheet.getRow(1), null,
    {
      pattern: "solid",
      fgColor: { argb: options?.fillColor ?? "ffffff" },
      bgColor: { argb: "ffffff" },
    }
  )
  if (options?.columnsNo) {
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
          color: { argb: options?.borderColor ?? "000000" },
        }
      )
    }
  }
}

const excelBeautify = {
  styleHeaders,
  createOuterBorder,
  fillCellWithPattern,
  fillCellsWithPattern,
  changeFontProperties,
  cellAlignmentForAllColumns,
}

export default excelBeautify