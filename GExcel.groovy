@Grab(group='org.apache.poi', module='poi', version='3.5-beta3')
import org.apache.poi.hssf.usermodel.*
import org.apache.poi.poifs.filesystem.*

class GExcel {
    static { setupMetaClass() }

    private static setupMetaClass() {
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
            asString { delegate.stringCellValue }
            asType { Class type ->
                switch(type) {
                    case Integer: return delegate.asInt()
                    case Boolean: return delegate.asBoolean()
                    case Date:    return delegate.asDate()
                    case String:  return delegate.asString()
                }
            }
        }
    }
    private static convertRowLabelToNumber(ascii) {
        final origin = ('A' as char) as int
        final radix = 26
        def num = 0
        ascii.toUpperCase().reverse().eachWithIndex { ch, i ->
            def delta = ((ch as char) as int) - origin + 1
            num += delta * (radix**i)
        }
        num - 1 // convert for "index" which starts from 0
    }
    // "A8" -> A -> 0
    private static colIndex(expr) {
        def matcher = (expr =~ /([a-zA-Z]+)([0-9]+)/)
        convertRowLabelToNumber(matcher[0][1])
    }
    // "A8" -> 8 -> 7
    private static rowIndex(expr) {
        def matcher = (expr =~ /([a-zA-Z]+)([0-9]+)/)
        (matcher[0][2] as int) - 1
    }

    static load(String file) { load(new File(file)) }
    static load(File file) { load(new FileInputStream(file)) }
    static load(InputStream is) { new HSSFWorkbook(new POIFSFileSystem(is)) }
}
