import jxl.*;
import jxl.write.*;

static void main(String[] args) {
    File txtFile = new File('input.txt')

    if (checkIfFileExistsAndNotEmpty(txtFile) == 1){
        createExcelFileAndSheet(txtFile)
    }
}

def createExcelFileAndSheet(txtFile) {
    // Create a workbook
    WritableWorkbook workbook = Workbook.createWorkbook(new File("output.xls"))
    WritableSheet sheet = workbook.createSheet("Translation", 0)

    // Names of the column labels
    println("file language: ")
    def inputString = System.in.newReader().readLine()
    println("translation language: ")
    def translationLanguage = System.in.newReader().readLine()

    // Add labels in the first row
    sheet.addCell(new Label(0, 0, "Label"))
    sheet.addCell(new Label(1, 0, inputString))
    sheet.addCell(new Label(2, 0, translationLanguage))

    //Separate text from txt file and fill the generated Excel file with data
    separateTextAndFillTheExcelFile(txtFile, sheet)

    // Write the data and close the workbook
    workbook.write()
    workbook.close()

    println ("Excel file generated.")
}

def separateTextAndFillTheExcelFile(txtFile, sheet) {
    Map<String, String> dataMap = [:]

    // this script helps to add together label info and added translation info if there is any
    // script in D1 field (have to drag it down)
    sheet.addCell(new Label(3, 0, "=IF(C1<>\"\",CONCAT(A1,\"=\",C1),\"\")"))

    def row = 0

    txtFile.eachLine { line ->
        row++
        def parts = line.split("=")
        if (parts.length == 2) {
            def column = 0

            sheet.addCell(new Label(column, row, parts[0].trim()))
            column++
            sheet.addCell(new Label(column, row, parts[1].trim()))

        }
    }
    println ("Data added in excel file.")
}

def checkIfFileExistsAndNotEmpty(txtFile){
    if (txtFile.exists() == true && txtFile.length() > 0){
        return 1
    }
    else if (txtFile.exists() == false) {
        println("Error 404 - file not found.")
        return 0}
    else if(txtFile.exists() == false && txtFile.length() == 0) {
        println("File is empty.")
        return 0}
    else {
        println("Unknown error")
        return 0
    }
}
