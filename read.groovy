@Grab(group='org.apache.poi', module='poi', version='3.5-beta3')
import org.apache.poi.hssf.usermodel.*
import org.apache.poi.poifs.filesystem.*

//---------------------------------------------------------
def ascii2num = { ascii ->
    def num = 0
    ascii.reverse().toUpperCase().toCharArray().eachWithIndex { ch, i ->
        final maxDelta = 25
        def deltaFromA = (ch as int) - 65
        num += deltaFromA * (i*maxDelta)
    }
    num
}

// "A8" -> 0
def colIndex = { expr ->
    def matcher = (expr =~ /([a-zA-Z]+)([0-9]+)/)
    println matcher[0][1]
    println ascii2num(matcher[0][1])
    ascii2num(matcher[0][1])
}
// "A8" -> 7
def rowIndex = { expr ->
    def matcher = (expr =~ /([a-zA-Z]+)([0-9]+)/)
    (matcher[0][2] as int) - 1
}

// staticではdefineが使えない(!?)ので。
HSSFWorkbook.metaClass.static.load = { String file -> load(new File(file)) }
HSSFWorkbook.metaClass.static.load = { File file -> load(new FileInputStream(file)) }
HSSFWorkbook.metaClass.static.load = { InputStream is -> new HSSFWorkbook(new POIFSFileSystem(is)) }

HSSFWorkbook.metaClass.define {
    getAt { idx -> delegate.getSheetAt(idx) }
    getProperty { name -> delegate.getSheet(name) }
}
HSSFSheet.metaClass.define {
    getProperty { name -> delegate.getRow(rowIndex(name))?.getCell(colIndex(name)) }
    setProperty { name, value -> delegate.getRow(rowIndex(name))?.getCell(colIndex(name)).setCellValue(value) }
}
HSSFCell.metaClass.define {
    asInt { delegate.numericCellValue.intValue() }
    asDate { delegate.dateCellValue }
    asBoolean { delegate.booleanCellValue }
}
class GExcel {
    static load(String file) { load(new File(file)) }
    static load(File file) { load(new FileInputStream(file)) }
    static load(InputStream is) { new HSSFWorkbook(new POIFSFileSystem(is)) }
}
//---------------------------------------------------------

//def book = HSSFWorkbook.load(args[0])
def book = GExcel.load(args[0])

def sheet = book[0]

println sheet["A1"]

println sheet.A1
sheet.A1 = "MODIFIED"
println sheet.A1

println sheet.A2
println sheet.AA1
println sheet.BC1
println sheet.ZZ1
println sheet.ZZZ1

//println sheet[A1..B5]
//sheet[A1..B5].filledBy 0
//println sheet[A1..B5]

def sheet3 = book.Sheet3 // by sheet name

println sheet3.A1
println sheet3.A2
println sheet3.A3.asInt()
println sheet3.A4.asDate()
println sheet3.A5.asBoolean()
println sheet3.A6.asBoolean()
println sheet3.B1
println sheet3.B2

