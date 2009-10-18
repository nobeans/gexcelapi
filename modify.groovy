@Grab(group='org.apache.poi', module='poi', version='3.5-beta3')
import org.apache.poi.hssf.usermodel.*
import org.apache.poi.poifs.filesystem.*

def inputFile = new File(args[0])
def book = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(inputFile)))
def sheet1 = book.getSheetAt(0) 
def sheet2 = book.getSheetAt(1) 
def sheet3 = book.getSheet("Sheet3") 

def sheet = sheet1 

def cell = { row, col ->
    sheet.getRow(row)?.getCell((short) col)
}

println cell(0, 0).stringCellValue 
println cell(1, 0).stringCellValue 
println cell(2, 0).numericCellValue.intValue() 
println cell(3, 0).dateCellValue 
println cell(4, 0).booleanCellValue 

cell(0, 0).setCellValue("Modified_A1") 
cell(1, 0).setCellValue("•ÏX‚µ‚½_A2") 
cell(2, 0).setCellValue(7890) 
cell(3, 0).setCellValue(new Date()) 
cell(4, 0).setCellValue(false) 

println cell(0, 0).stringCellValue 
println cell(1, 0).stringCellValue 
println cell(2, 0).numericCellValue.intValue() 
println cell(3, 0).dateCellValue 
println cell(4, 0).booleanCellValue 

def outputFile = new File(args[1])
outputFile.withOutputStream { out ->
    book.write(out)
}

