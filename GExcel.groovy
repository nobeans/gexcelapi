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
        }
    }
    private static ascii2num(ascii) {
        def num = 0
        ascii.reverse().toUpperCase().toCharArray().eachWithIndex { ch, i ->
            final maxDelta = 25
            def deltaFromA = (ch as int) - 65
            num += deltaFromA * (i*maxDelta)
        }
        num
    }
    // "A8" -> A -> 0
    private static colIndex(expr) {
        def matcher = (expr =~ /([a-zA-Z]+)([0-9]+)/)
        println matcher[0][1]
        println ascii2num(matcher[0][1])
        ascii2num(matcher[0][1])
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
