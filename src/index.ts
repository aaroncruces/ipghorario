import { WorkBook, readFile, utils } from "xlsx"
import { Profesor } from "./types/Profesor"
import { formatName } from "./utils/string_utils"
import { stripRutSpecialCharacters } from "./utils/rut_utils"
console.log("..........................................")
const main = () => {
    const workbookIPG = readFile("xlsx_samples/IPG_Datos.xlsx")
    const sheetProfesores = extractProfesorListFromWorkbook(workbookIPG)

    //console.log(JSON.stringify(workbookIPG.Sheets["Profesores"], null, 2))
    //console.log(JSON.stringify(sheetProfesores, null, 2))
    console.log(JSON.stringify(sheetProfesores, null, 2))
}

const extractProfesorListFromWorkbook = (workbook: WorkBook): Array<Profesor | undefined> => {
    const rawSheetProfesores = workbook.Sheets["Profesores"]
    const unprocessedSheetProfesores = utils.sheet_to_json(rawSheetProfesores)
    const sheetProfesores = processProfesoresFromSheet(unprocessedSheetProfesores)
    return sheetProfesores
}

const processProfesoresFromSheet = (unprocessedProfesorList: Array<any>): Array<Profesor | undefined> => {
    const reformattedProfesorList =
        unprocessedProfesorList.map(
            (unprocessedProfesor: any): Profesor | undefined =>
                processProfesorFromSheet(unprocessedProfesor)
        )
    const cleanProfesorList = reformattedProfesorList
    return cleanProfesorList
}

const processProfesorFromSheet = (unprocessedProfesor: any): Profesor | undefined => {
    //Rut must be a string on xlsx sheet
    let newRut: string | number = unprocessedProfesor["Rut"]
    if (!newRut || newRut == "")
        return undefined
    //TODO: REGEX RUT
    //TODO: GET NUMERICPART WITH REGEX
    //TODO: FORMAT RUT ON XLSX WRITING
    newRut = stripRutSpecialCharacters(newRut.toString())


    //Nombre must be a string on xlsx sheet
    let newNombre: string = unprocessedProfesor["Nombre"]
    if (!newNombre || newNombre == "")
        return undefined

    newNombre = formatName(newNombre)

    //Telefono shoud be a number on xlsx sheet
    let newTelefono: string | number = unprocessedProfesor["Telefono"]
    //TODO: CONVERT TO NUMBER

    //Must be String
    let newCorreoIpg: string = unprocessedProfesor["Correo IPG"]

    //Must be String
    let newCorreoPersonal: string = unprocessedProfesor["Correo Personal"]

    //
    let newComentario: string = unprocessedProfesor["Comentarios"]


    const newProfesor: Profesor = {
        nombre: newNombre,
        rut: newRut,
        telefono: newTelefono,
        correoipg: newCorreoIpg,
        correopersonal: newCorreoPersonal,
        comentarios: newComentario,
    }
    return newProfesor
}

main()
