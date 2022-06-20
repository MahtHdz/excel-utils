const genSheetRows = (data, mode="flexible") => {
  const fields = Object.keys(data)
  const rawData = Object.values(data)
  let rows = []
  let maxRows = 0
  let strictModeFlag = true

  if(mode === "flexible") {
    rawData.forEach(column => {
      if(column.length > maxRows) {
        maxRows = column.length
      }
    })
  }else if(mode === "strict") {
    rawData.forEach(column => {
      if (maxRows === 0) maxRows = column.length
      if(column.length !== maxRows) {
        strictModeFlag = false
      }
    })
  }
  if(mode === "flexible" || (mode === "strict" && strictModeFlag)) {
    for(let i = 0; i < maxRows; i++) {
      let row = {}
      fields.forEach((field, index) => {
        if(rawData[index][i] !== undefined)
          row[field] = rawData[index][i]
        else
          row[field] = null
      })
      rows.push(row)
    }
  }
  return rows
}

const genSheetRowsWithContext = (data, mode="flexible") => {
  const fields = Object.keys(data)
  const rows = genSheetRows(data, mode)
  return { fields, rows }
}

const excelDataProcessing = {
  genSheetRows,
  genSheetRowsWithContext
}

export default excelDataProcessing