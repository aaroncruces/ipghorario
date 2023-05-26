

export const formatName: (name: string) => string = (name: string) => {
    let formattedName = name.toLowerCase()
    formattedName = formattedName.replace(/à/g, "á").replace(/è/g, "é").replace(/ì/g, "í").replace(/ò/g, "ó").replace(/ù/g, "ú")
    formattedName = formattedName.replace(/[^a-zñáéíóúü´`¨\- ]/g, "")
    formattedName = formattedName.replace(/\s+/g, ' ').trimStart().toLowerCase()

    return formattedName
}

export const decorateName: (name: string) => string = (name: string) => {
    const separateWords = name.split(" ")
    const decoratedName = separateWords.reduce((
        (accumulated: string, currentWord: string, index: number) => {
            let separator = " "
            if (index === (separateWords.length - 1)) {
                separator = ""
            }
            return accumulated + capitalizeWord(currentWord) + separator
        }), "")
    return decoratedName
}

export const formatAlphanumeric: (name: string) => string = (name: string) => {
    let formattedName = name.toLowerCase()
    formattedName = formattedName.replace(/à/g, "á").replace(/è/g, "é").replace(/ì/g, "í").replace(/ò/g, "ó").replace(/ù/g, "ú")
    formattedName = formattedName.replace(/[^0-9a-zñáéíóúü´`¨\- ]/g, "")
    const separateWords = formattedName.replace(/\s+/g, ' ').trimStart().toLowerCase().split(" ")
    formattedName = separateWords.reduce((
        (accumulated: string, currentWord: string, index: number) => {
            let separator = " "
            if (index === (separateWords.length - 1)) {
                separator = ""
            }
            return accumulated + capitalizeWord(currentWord) + separator
        }), "")

    return formattedName
}

export const formatNumeric: (numericString: string) => string =
    (numericString: string) => numericString.replace(/[^0-9]/g, "");

export const capitalizeWord: (word: string) => string = (word: string) => word.length <= 1 ? word.toUpperCase() : word[0].toUpperCase() + word.substring(1)

export const formatEmail: (email: string) => string = (email: string) => {
    // https://en.wikipedia.org/wiki/Email_address#Local-part
    //multiple dots are not allowed
    const splitEmail = email.toLowerCase().replace(/\.\./g, ".").split("@")
    let localPart = splitEmail[0]
    //assuming unquoted local part
    // if needed to use a more general localpart:
    //localPart = localPart.replace(/[^a-z0-9!#$%&'*+-/=?^_`{|}~]/g, "")
    // assuming hotmail-gmail like email:
    localPart = localPart.replace(/[^a-z0-9-_.]/g, "")

    //only localpart
    if (splitEmail.length <= 1) return localPart;

    splitEmail.shift()
    //domainpart only has a single "@"
    let domainPart = splitEmail.join("")
    domainPart = domainPart.replace(/[^a-z0-9.]/g, "")
    //domainpart cannot start with "-"
    if (domainPart[0] === "-") {
        if (domainPart.length === 1) domainPart = ""
        else domainPart = domainPart.substring(1);
    }

    return localPart + "@" + domainPart
}

export const checkEmail: (email: string) => string = (email: string) => {
    if (!email) return ""
    if (email.length < 1) return ""

    const splitAT = email.split("@")
    if (splitAT.length !== 2) return "Email debe tener arroba '@'"

    if (splitAT[0] === "") return "Email debe tener nombre de usuario 'usuario@'"

    if (splitAT[0][0] === ".") return "Email no puede empezar con punto '.'"
    if (splitAT[0][0] === "-") return "Email no puede empezar con guion '-'"
    if (splitAT[0][0] === "_") return "Email no puede empezar con guion bajo '_'"

    const domainSplitDOT = splitAT[1].split(".")
    if (domainSplitDOT.length <= 1) return "Email debe tener dominio de la forma '@sitio.dom'"

    if (domainSplitDOT[0] === "") return "Email debe tener dominio de la forma '@sitio.dom'"

    const lastWord = domainSplitDOT[domainSplitDOT.length - 1]
    if (lastWord.length < 1) return "Email no puede terminar con punto '.'"

    if (email.toLowerCase().trim() !== formatEmail(email).toLowerCase().trim())
        return "Email contiene caracteres invalidos"

    return ""
}

export const stripIntegerStringSpecialCharacters: (numericString: string) => string =
    (numericString: string) => numericString.replace(/[^0-9]/g, "")

//todo: check  phone number rules on Chile
export const decoratePhoneNumber: (numericString: string) => string =
    (numericString: string) => numericString


//todo: check  phone number rules on Chile
export const decorateAddressNumber: (numericString: string) => string =
    (numericString: string) => numericString

