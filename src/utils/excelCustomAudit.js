const fixNumberLength = (number, length) => {
    const strNumber = number.toString()
    if(strNumber.length !== length){
        let missingZeros = ""
        for (let i = 0; i < length; i++)
            missingZeros += "0"
        return missingZeros + strNumber
    }
}

export const doNRectifications = (PU501, PUF701, PUA701) => {
    //const generalLengthArray_701 = PUA701.length
    const rectifications = []
    let stopFlag = false
    PU501.forEach(PU => {
        let rectificationArray = []
        let currentId = PU
        while(!stopFlag){
            const tmpIndex = PUF701.indexOf(currentId)
            if(tmpIndex !== -1) {
                rectificationArray.push(currentId)
                rectificationArray.push(PUA701[tmpIndex])
                currentId = PUA701[tmpIndex]
            } else stopFlag = true
        }
        stopFlag = false
        if(rectificationArray.length > 0) rectifications.push([ ...new Set(rectificationArray) ])
    })
    return rectifications
}

const genUniqueId = (CSA, PA, NP) => { 
    const CSARegex = /^[0-9]{3}$/
    const PARegex = /^[0-9]{4}$/
    const NPRegex = /^[0-9]{7}$/
    const IdRegex = /^[0-9]{3}-[0-9]{4}-[0-9]{7}$/
    let strCSA = ""
    let strPA = ""
    let strNP = ""

    if(CSA instanceof Number) strCSA = fixNumberLength(CSA, 3)
    else if(CSA instanceof String) strCSA = CSA
    if(PA instanceof Number) strPA = fixNumberLength(PA, 4)
    else if(PA instanceof String) strPA = PA
    if(NP instanceof Number) strNP = fixNumberLength(NP, 7)
    else if(NP instanceof String) strNP = NP
    let id = CSARegex.test(strCSA) && PARegex.test(strPA) && NPRegex.test(strNP) ? `${strCSA}-${strPA}-${strNP}` : ''
    if(!IdRegex.test(id))
        id = '000-0000-0000000'
    return id
}

export const execDataStageAudit = (data) => {
    
}