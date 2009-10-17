// usage: groovy modify.groovy read.xls result.xls

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
// セルの書き換え
// HSSFCellの特定は読み込みと同じであるため、簡単のためヘルパメソッドを利用する
def cell = { row, col ->
    sheet.getRow(row)?.getCell((short) col)
}

// 事前の値確認
println cell(0, 0).stringCellValue // A1
println cell(1, 0).stringCellValue // A2
println cell(2, 0).numericCellValue.intValue() // A3 (double->int)
println cell(3, 0).dateCellValue // A4
println cell(4, 0).booleanCellValue // A5

// 書き換え
cell(0, 0).setCellValue("Modified_A1") // A1
cell(1, 0).setCellValue("変更した_A2") // A2
cell(2, 0).setCellValue(7890) // A3
cell(3, 0).setCellValue(new Date()) // A4
cell(4, 0).setCellValue(false) // A5

// 事後の値確認
println cell(0, 0).stringCellValue // A1
println cell(1, 0).stringCellValue // A2
println cell(2, 0).numericCellValue.intValue() // A3 (double->int)
println cell(3, 0).dateCellValue // A4
println cell(4, 0).booleanCellValue // A5

// ----------------------------
// 新規Excelファイルへの出力
def outputFile = new File(args[1])
outputFile.withOutputStream { out ->
    book.write(out)
}

