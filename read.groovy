@Grab(group='org.apache.poi', module='poi', version='3.5-beta3')
import org.apache.poi.hssf.usermodel.*
import org.apache.poi.poifs.filesystem.*

def inputFile = new File(args[0])
def book = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(inputFile)))
def sheet1 = book.getSheetAt(0) 
def sheet2 = book.getSheetAt(1) 
def sheet3 = book.getSheet("Sheet3") 

def sheet = sheet1 

println sheet.getRow(0)?.getCell((short) 0).stringCellValue 
println sheet.getRow(1)?.getCell((short) 0).stringCellValue 
println sheet.getRow(2)?.getCell((short) 0).numericCellValue.intValue() 
println sheet.getRow(3)?.getCell((short) 0).dateCellValue 
println sheet.getRow(4)?.getCell((short) 0).booleanCellValue 
println sheet.getRow(5)?.getCell((short) 0).booleanCellValue 
println sheet.getRow(0)?.getCell((short) 0).stringCellValue 
println sheet.getRow(0)?.getCell((short) 1).stringCellValue 

def cell = { row, col ->
    sheet.getRow(row)?.getCell((short) col)
}
println cell(0, 0).stringCellValue 
println cell(1, 0).stringCellValue 
println cell(2, 0).numericCellValue.intValue() 
println cell(3, 0).dateCellValue 
println cell(4, 0).booleanCellValue 
println cell(5, 0).booleanCellValue 
println cell(0, 0).stringCellValue 
println cell(0, 1).stringCellValue 

