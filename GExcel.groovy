@Grab(group='org.apache.poi', module='poi', version='3.5-beta3')
import org.apache.poi.hssf.usermodel.*
import org.apache.poi.poifs.filesystem.*
import org.apache.poi.ss.usermodel.*
import static Util.*

class GExcel {
    static { setupMetaClass() }

    private static setupMetaClass() {
        HSSFWorkbook.metaClass.define {
            getAt { int idx -> delegate.getSheetAt(idx) }
            getProperty { String name -> delegate.getSheet(name) }
        }
        HSSFSheet.metaClass.define {
            getProperty { name ->
                if (name == "rows") return rows()
                if (name ==~ /_\d+/) return delegate.getRow(rowIndex(name))
                try {
                    delegate.getRow(rowIndex(name))?.getCell(colIndex(name))
                } catch (IndexOutOfBoundsException e) {
                    return null
                }
            }
            setProperty { name, value ->
                try {
                    delegate.getRow(rowIndex(name))?.getCell(colIndex(name)).setCellValue(value)
                } catch (IndexOutOfBoundsException e) {
                    return null
                }
            }
            rows { continueCondition = {true} ->
                def rows = []
                for (def row : delegate) {
                    if (!continueCondition.call(row)) break
                    rows << row
                }
                rows
            }
        }
        HSSFRow.metaClass.define {
            getAt { int idx -> delegate.getCell(idx) }
            getProperty { name ->
                if (name ==~ /[a-zA-Z]+_/) return delegate.getCell(colIndex(name))
                null
            }
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
            toString { delegate.value as String }
        }
    }

    static load(String file) { load(new File(file)) }
    static load(File file) { load(new FileInputStream(file)) }
    static load(InputStream is) { new HSSFWorkbook(new POIFSFileSystem(is)) }
}

