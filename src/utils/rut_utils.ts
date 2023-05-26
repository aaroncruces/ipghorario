//
/**
 * Assumes rut with dv
 * takes a string with non escential rut-characters (like 12.345 684-K)
 * to a rut with rut with de bare minimum 
 * @param formattedRut 12.3450 684-K
 * @returns 12345684K
 */
export const stripRutSpecialCharacters: (formattedRut: string) => string =
    (formattedRut: string) => formattedRut.replace(/[^0-9kK]/g, "").replace(/[k]/g, "K");

export const decorateRut: (formattedRut: string) => string =
    (formattedRut: string) => {
        const strippedRut = stripRutSpecialCharacters(formattedRut)
        if (strippedRut.length === 1) {
            return formattedRut
        }
        if (strippedRut.length > 1) {
            return separateNumberStringWithDots(strippedRut.slice(0, (strippedRut.length - 1))) + "-" + strippedRut.slice(strippedRut.length - 1)
        }
        return ""
    }

/**
 * takes a number and like "12345678" and returns a formatted number (12.345.678)
 * credits to: https://stackoverflow.com/questions/2901102/how-to-format-a-number-with-commas-as-thousands-separators
 * TODO: fix, is a bit slow
 * @param digits 
 * @returns 
 */
const separateNumberStringWithDots: (digits: string) => string =
    (digits: string) =>
        digits.replace(/\B(?=(\d{3})+(?!\d))/g, ".");


/**
 * checks a RUT in the form "182360260"
 * @param rutWithDV 
 * @returns errorMessage
 */
export const checkStructure_rutWithDv: (rutWithDv: string) => string =
    (rutWithDV: string) => {

        if (rutWithDV === undefined || rutWithDV === null)
            return "Rut invalido";

        if (rutWithDV === "")
            return "Rut es obligatorio";

        if (rutWithDV.length === 1) {
            if (isNaN(Number(rutWithDV[0])))
                return "Rut debe ser numerico";
            return "Rut no puede ser de 1 solo digito"
        }

        return ""
    }

export const checkDV_rutWithDV: (rutWithDV: string) => string = (rutWithDV: string) => {

    const numericPart = rutWithDV.slice(0, (rutWithDV.length - 1))
    if (isNaN(Number(numericPart)))
        return "El rut debe ser numerico";

    const dvPart = rutWithDV[rutWithDV.length - 1]
    const calculatedDV = calculate_dv(Number(numericPart))

    if (dvPart !== calculatedDV)
        return "Digito Verificador Incorrecto"

    return "";
}

export const calculate_dv: (rutWithDv: number) => string = (rutWithoutDV: number) => {
    var M = 0, S = 1;
    for (; rutWithoutDV; rutWithoutDV = Math.floor(rutWithoutDV / 10))
        S = (S + rutWithoutDV % 10 * (9 - M++ % 6)) % 11;
    const result = S ? S - 1 : 'K';
    return result.toString()
}
