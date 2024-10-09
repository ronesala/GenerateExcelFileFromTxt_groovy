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

    // Name the columns
    addLabelsToColumns(sheet)
    //Separate text from txt file and fill the generated Excel file with data
    separateTextAndFillTheExcelCells(txtFile, sheet)

    // Write the data and close the workbook
    workbook.write()
    workbook.close()

    println ("Excel file generated.")
}

def addLabelsToColumns(sheet){
    // Names of the column labels
    println("file language: ")
    def inputString = System.in.newReader().readLine()
    println("translation language: ")
    def translationLanguage = System.in.newReader().readLine()

    // Add labels in the first row
    sheet.addCell(new Label(0, 0, "Label"))
    sheet.addCell(new Label(1, 0, inputString))
    sheet.addCell(new Label(2, 0, translationLanguage))
    sheet.addCell(new Label(3, 0, "Label = " + translationLanguage))
}

def separateTextAndFillTheExcelCells(txtFile, sheet) {
    Map<String, String> dataMap = [:]

    def row = 0
    def checkRow = 1

    txtFile.eachLine { line ->
        row++
        checkRow++

        def parts = line.split("=")
        if (parts.length == 2) {
            sheet.addCell(new Label(0, row, parts[0].trim()))
            sheet.addCell(new Label(1, row, parts[1].trim()))
            sheet.addCell(new Label(2, row, ""))
            sheet.addCell(new Formula(3, row, "IF(C${checkRow}<>\"\",CONCATENATE(A${checkRow},\"=\",C${checkRow}),\"\")"))
        }
    }
    println ("Data added in excel file.")
}

def checkIfFileExistsAndNotEmpty(txtFile){
    if (txtFile.exists() == true && txtFile.length() > 0){
        return 1
    }

    if (txtFile.exists() == false) {
        println("Error 404 - file not found.")
        return 0
    }

    if (txtFile.exists() == false && txtFile.length() == 0) {
        println("File is empty.")
        return 0
    }

    println("Unknown error")
    return 0
}
