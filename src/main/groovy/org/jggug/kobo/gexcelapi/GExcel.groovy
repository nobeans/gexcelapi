/*
 * Copyright 2009 the original author or authors.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package org.jggug.kobo.gexcelapi

import org.apache.poi.ss.usermodel.*
import org.jggug.kobo.gexcelapi.CellLabelUtils as CLU
import java.lang.IndexOutOfBoundsException as IOOBEx

class GExcel {

    static {
        expandWorkbook()
        expandSheet()
        expandRow()
        expandCell()
    }

    private static expandWorkbook() {
        Workbook.metaClass.define {
            getAt { int idx -> delegate.getSheetAt(idx) }
            getProperty { String name -> delegate.getSheet(name) }
        }
    }

    private static expandSequentialCellRangeInstanceAsWildcardForRow(SequentialCellRange range) {
        range.metaClass.getProperty = { name ->
            if (name ==~ /[a-zA-Z]+_/) { // wildcard for column
                int columnIndex = CLU.columnIndex(name)
                return range[columnIndex]
            }
            return delegate[name]
        }
    }

    private static expandSequentialCellRangeInstanceAsWildcardForColumn(SequentialCellRange range) {
        range.metaClass.getProperty = { name ->
            if (name ==~ /_\d+/) { // wildcard for row
                int rowIndex = CLU.rowIndex(name)
                return range[rowIndex]
            }
            return delegate[name]
        }
    }

    private static expandSheet() {
        def methods = {
            getProperty { name ->
                if (name == "rows") { return rows() }
                if (name ==~ /_\d+/) { // wildcard for row
                    int rowIndex = CLU.rowIndex(name)
                    def row = delegate.getRow(CLU.rowIndex(name))
                    if (!row) {
                        row = delegate.createRow(rowIndex)
                        row.createCell(0)
                    }
                    def range = new SequentialCellRange(delegate, rowIndex, row.getFirstCellNum(), rowIndex, row.getLastCellNum() - 1)
                    expandSequentialCellRangeInstanceAsWildcardForRow(range)
                    return range
                }
                if (name ==~ /[a-zA-Z]+_/) { // wildcard for column
                    int columnIndex = CLU.columnIndex(name)
                    def range = new SequentialCellRange(delegate, delegate.getFirstRowNum(), columnIndex, delegate.getLastRowNum(), columnIndex)
                    expandSequentialCellRangeInstanceAsWildcardForColumn(range)
                    return range
                }
                if (name ==~ /[a-zA-Z]+\d+/) { // a specified cell
                    int rowIndex = CLU.rowIndex(name)
                    def row = delegate.getRow(CLU.rowIndex(name))
                    if (!row) {
                        row = delegate.createRow(rowIndex)
                        row.createCell(0)
                    }
                    int columnIndex = CLU.columnIndex(name)
                    try {
                        return row.getCell(CLU.columnIndex(name))
                    } catch (IOOBEx e) {
                        return row.createCell(columnIndex)
                    }
                }
                if (name ==~ /([a-zA-Z]+\d+)_([a-zA-Z]+\d+)/) { // cells in a specified rectangle
                    def token = name.split("_")
                    return new RectangleCellRange(delegate, token[0], token[1])
                }
                return delegate[name]
            }
            setProperty { name, value ->
                if (name ==~ /[a-zA-Z]+\d+/) {
                    int rowIndex = CLU.rowIndex(name)
                    def row = delegate.getRow(CLU.rowIndex(name))
                    if (!row) {
                        row = delegate.createRow(rowIndex)
                        row.createCell(0)
                    }
                    int columnIndex = CLU.columnIndex(name)
                    try {
                        return row.getCell(CLU.columnIndex(name)).setCellValue(value)
                    } catch (IOOBEx e) {
                        return row.createCell(columnIndex).setCellValue(value)
                    }
                }
                delegate[name] = value
            }
            rows { delegate?.findAll{true} }
            validate { delegate.rows.every { row -> row.validate() } }
        }
        Sheet.metaClass.define methods
    }

    private static expandRow() {
        def methods = {
            getAt { int idx -> delegate.getCell(idx) }
            getProperty { name ->
                if (name ==~ /[a-zA-Z]+_/) {
                    return delegate.getCell(CLU.columnIndex(name))
                }
                delegate[name]
            }
            getLabel { CLU.rowLabel(delegate.getRowNum()) }
            validate { delegate.every { cell -> cell.validate() } }
        }
        Row.metaClass.define methods
    }

    private static expandCell() {
        Cell.metaClass.__validators__ = null
        Cell.metaClass.define {
            isStringType  { delegate.cellType == Cell.CELL_TYPE_STRING }
            isNumericType { delegate.cellType == Cell.CELL_TYPE_NUMERIC }
            isBooleanType { delegate.cellType == Cell.CELL_TYPE_BOOLEAN }
            getValue {
                // implicitly accessing value by appropriate type
                switch(delegate.cellType) {
                    case Cell.CELL_TYPE_NUMERIC: return delegate.numericCellValue
                    case Cell.CELL_TYPE_STRING:  return delegate.stringCellValue
                    case Cell.CELL_TYPE_FORMULA: return delegate.cellFormula
                    case Cell.CELL_TYPE_BLANK:   return null
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
            getLabel { CLU.cellLabel(delegate.rowIndex, delegate.columnIndex) }
            getColumnLabel { CLU.columnLabel(delegate.columnIndex) }
            getRowLabel { CLU.rowLabel(delegate.rowIndex) }
            clearValidators {
                delegate.__validators__ = []
            }
            getValidators {
                if (delegate.__validators__ == null) { delegate.clearValidators() }
                delegate.__validators__
            }
            addValidator { Closure validator ->
                delegate.validators << validator
            }
            setValidator { Closure validator ->
                delegate.clearValidators()
                delegate.validators << validator
            }
            validate {
                delegate.validators.every { validator -> validator.call(delegate) }
            }
        }
    }

    static open(String file) { open(new File(file)) }
    static open(File file) { open(new FileInputStream(file)) }
    static open(InputStream is) { WorkbookFactory.create(is) }
}

