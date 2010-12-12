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

import org.apache.poi.ss.usermodel.Sheet
import org.jggug.kobo.gexcelapi.CellLabelUtils as CLU

class CellRange implements Range {

    @Delegate
    private List<List> list

    final Sheet sheet
    final int beginRow, beginColumn, endRow, endColumn

    static CellRange newSequentialCellRange(Sheet sheet, int beginRow, int beginColumn, int endRow, int endColumn) {
        return new CellRange(sheet, beginRow, beginColumn, endRow, endColumn, true)
    }

    static CellRange newSequentialCellRange(Sheet sheet, String beginCellLabel, String endCellLabel) {
        return new CellRange(sheet, beginCellLabel, endCellLabel, true)
    }

    static CellRange newRectangleCellRange(Sheet sheet, int beginRow, int beginColumn, int endRow, int endColumn) {
        return new CellRange(sheet, beginRow, beginColumn, endRow, endColumn, false)
    }

    static CellRange newRectangleCellRange(Sheet sheet, String beginCellLabel, String endCellLabel) {
        return new CellRange(sheet, beginCellLabel, endCellLabel, false)
    }

    private CellRange(Sheet sheet, int beginRow, int beginColumn, int endRow, int endColumn, boolean flatten=false) {
        this.sheet = sheet
        this.beginRow = beginRow
        this.beginColumn = beginColumn
        this.endRow = endRow
        this.endColumn = endColumn
        this.list = new CellLabelIterator(beginRow, beginColumn, endRow, endColumn).collect { row ->
            row.collect { label ->
                sheet[label] ?: sheet.createRow(CLU.rowIndex(label)).createCell(CLU.columnIndex(label))
            }
        }
        if (flatten) {
            this.list = list.flatten() // to 1 dimension
        }
    }

    private CellRange(Sheet sheet, String beginCellLabel, String endCellLabel, boolean flatten=false) {
        this(sheet, CLU.rowIndex(beginCellLabel), CLU.columnIndex(beginCellLabel), CLU.rowIndex(endCellLabel), CLU.columnIndex(endCellLabel), flatten)
    }

    def getProperty(String name) {
        if (name ==~ /[a-zA-Z]+_/) { // wildcard for column
            int columnIndex = CLU.columnIndex(name)
            return this[columnIndex]
        }
        if (name ==~ /_\d+/) { // wildcard for row
            int rowIndex = CLU.rowIndex(name)
            return this[rowIndex]
        }
        return delegate[name]
    }

    boolean validate() {
        list.every { row ->
            row.every { cell ->
                cell?.validate()
            }
        }
    }

    @Override
    boolean containsWithinBounds(Object o) {
        list.contains(o)
    }

    @Override
    Comparable getFrom() {
        list.first()
    }

    @Override
    Comparable getTo() {
        list.tail()
    }

    @Override
    String inspect() {
        "#$list"
    }

    @Override
    boolean isReverse() {
        false // fixed
    }

    @Override
    List step(int step) {
        throw new UnsupportedOperationException("not implemented")
    }

    @Override
    void step(int step, Closure closure) {
        throw new UnsupportedOperationException("not implemented")
    }
}
