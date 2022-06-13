import excelDataExtractor from "../utils/excelDataExtractor"

const getDataFromFile = async (req, res) => {
    const { filepath, sheetName, fields } = req.body
    const data = await excelDataExtractor.getData(filepath, sheetName, fields)
    res.send(data)
}

const excelController = {
    getDataFromFile
}

export default excelController