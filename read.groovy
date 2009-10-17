// usage: groovy read.groovy read.xls

// ----------------------------
// Grapeによるライブラリ取得
@Grab(group='org.apache.poi', module='poi', version='3.5-beta3')
import org.apache.poi.hssf.usermodel.*
import org.apache.poi.poifs.filesystem.*
//groovy.grape.Grape.grab(group:'org.apache.poi', module:'poi', version:'3.0.2-FINAL')
//@Grab('org.apache.poi:poi:3.0.2-FINAL') // only for Groovy1.7

// ----------------------------
// Excelファイルの読み込み
def inputFile = new File(args[0])
def book = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(inputFile)))
def sheet1 = book.getSheetAt(0) // 第1シート
def sheet2 = book.getSheetAt(1) // 第2シート
def sheet3 = book.getSheet("Sheet3") // シート名で指定も可能

def sheet = sheet1 // 以降で使うシートを選択
//def sheet = sheet2
//def sheet = sheet3

// ----------------------------
// セルの参照
println sheet.getRow(0)?.getCell((short) 0).stringCellValue // A1
println sheet.getRow(1)?.getCell((short) 0).stringCellValue // A2
println sheet.getRow(2)?.getCell((short) 0).numericCellValue.intValue() // A3 (double->int)
println sheet.getRow(3)?.getCell((short) 0).dateCellValue // A4
println sheet.getRow(4)?.getCell((short) 0).booleanCellValue // A5
println sheet.getRow(5)?.getCell((short) 0).booleanCellValue // A6
println sheet.getRow(0)?.getCell((short) 0).stringCellValue // B1
println sheet.getRow(0)?.getCell((short) 1).stringCellValue // B2

// （特に意味はないが）groovy風にクロージャで定義してみる
def cell = { row, col ->
    sheet.getRow(row)?.getCell((short) col)
}
println cell(0, 0).stringCellValue // A1
println cell(1, 0).stringCellValue // A2
println cell(2, 0).numericCellValue.intValue() // A3 (double->int)
println cell(3, 0).dateCellValue // A4
println cell(4, 0).booleanCellValue // A5
println cell(5, 0).booleanCellValue // A6
println cell(0, 0).stringCellValue // B1
println cell(0, 1).stringCellValue // B2

