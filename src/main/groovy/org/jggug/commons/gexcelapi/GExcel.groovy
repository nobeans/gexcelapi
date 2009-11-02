package org.jggug.commons.gexcelapi

import org.apache.poi.hssf.usermodel.*
import org.apache.poi.poifs.filesystem.*
import org.apache.poi.ss.usermodel.*
import static Util.*
import java.lang.IndexOutOfBoundsException as IOOBEx

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
                if (name ==~ /[a-zA-Z]+\d+/) {
                    try { delegate.getRow(rowIndex(name))?.getCell(colIndex(name)) } catch (IOOBEx e) { return null }
                }
            }
            setProperty { name, value ->
                if (name ==~ /[a-zA-Z]+\d+/) {
                    try { delegate.getRow(rowIndex(name))?.getCell(colIndex(name)).setCellValue(value) } catch (IOOBEx e) { return null }
                }
            }
            rows { delegate?.findAll{true} }
        }
        HSSFRow.metaClass.define {
            getAt { int idx -> delegate.getCell(idx) }
            getProperty { name ->
                if (name ==~ /[a-zA-Z]+_/) return delegate.getCell(colIndex(name))
                null
            }
        }
        HSSFCell.metaClass.define {
            isStringType  { delegate.cellType == Cell.CELL_TYPE_STRING }
            isNumericType { delegate.cellType == Cell.CELL_TYPE_NUMERIC }
            isBooleanType { delegate.cellType == Cell.CELL_TYPE_BOOLEAN }
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
                // explicitly accessing value by appropriate type
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

