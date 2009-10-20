@Grab(group='org.apache.poi', module='poi', version='3.5-beta3')
import org.apache.poi.hssf.usermodel.*
import org.apache.poi.poifs.filesystem.*
import org.apache.poi.ss.usermodel.*

class GExcel {
    static { setupMetaClass() }

    private static setupMetaClass() {
        // maybe cannot use "define" for static methods
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
            getValue {
                // implicitly accessing value by appropriate type
                switch(delegate.cellType) {
                    case Cell.CELL_TYPE_STRING:  return delegate.stringCellValue
                    case Cell.CELL_TYPE_NUMERIC: return delegate.numericCellValue
                    case Cell.CELL_TYPE_BOOLEAN: return delegate.booleanCellValue
                    default: throw new RuntimeException("unsupported cell type: ${delegate.cellType}")
                }
            }
            setValue { value -> delegate.setCellValue(value) }
            leftShift { value -> delegate.setCellValue(value) }
            asType { Class type ->
                switch(type) {
                    case Double:  return delegate.numericCellValue
                    case Integer: return delegate.numericCellValue.intValue()
                    case Boolean: return delegate.booleanCellValue
                    case Date:    return delegate.dateCellValue
                    case String:  return delegate.stringCellValue
                    default: throw new RuntimeException("unsupported cell type: ${delegate.cellType}")
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
