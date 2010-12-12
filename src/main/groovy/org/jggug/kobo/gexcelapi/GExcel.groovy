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

    private static getRowFromSheetByLabel(sheet, label) {
        int rowIndex = CLU.rowIndex(label)
        def row = sheet.getRow(rowIndex)
        if (!row) {
            row = sheet.createRow(rowIndex)
            row.createCell(0, Cell.CELL_TYPE_BLANK)
        }
        return row
    }

    private static getCellFromRowByLabel(row, label) {
        int columnIndex = CLU.columnIndex(label)
        def cell = row.getCell(columnIndex)
        if (!cell) {
            cell = row.createCell(columnIndex, Cell.CELL_TYPE_BLANK)
        }
        return cell
    }

    private static expandSheet() {
        def methods = {
            getProperty { name ->
                if (name == "rows") { return rows() }
                if (name ==~ /_\d+/) { // wildcard for row
                    def row = getRowFromSheetByLabel(delegate, name)
                    def range = CellRange.newSequentialCellRange(delegate, row.rowNum, row.getFirstCellNum(), row.rowNum, row.getLastCellNum() - 1)
                    return range
                }
                if (name ==~ /[a-zA-Z]+_/) { // wildcard for column
                    int columnIndex = CLU.columnIndex(name)
                    def range = CellRange.newSequentialCellRange(delegate, delegate.getFirstRowNum(), columnIndex, delegate.getLastRowNum(), columnIndex)
                    return range
                }
                if (name ==~ /[a-zA-Z]+\d+/) { // a specified cell
                    def row = getRowFromSheetByLabel(delegate, name)
                    def cell = getCellFromRowByLabel(row, name)
                    return cell
                }
                if (name ==~ /([a-zA-Z]+\d+)_([a-zA-Z]+\d+)/) { // cells in a specified rectangle
                    def token = name.split("_")
                    return CellRange.newRectangleCellRange(delegate, token[0], token[1])
                }
                return delegate[name]
            }
            setProperty { name, value ->
                if (name ==~ /[a-zA-Z]+\d+/) {
                    def row = getRowFromSheetByLabel(delegate, name)
                    def cell = getCellFromRowByLabel(row, name)
                    cell.setCellValue(value)
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
                    def cell = getCellFromRowByLabel(delegate, name)
                    return cell
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

