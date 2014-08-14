/*
 * Copyright 2009 the original author or authors.
 *
 * Licensed under the Apache License, Version 2.0 (the "License")
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

class GExcel_RowFinder_Test extends GroovyTestCase {

    def sampleFile = "build/resources/test/GExcel_RowFinder_Test-input.xls"
    def sheet

    void setUp() {
        def book = GExcel.open(sampleFile)
        sheet = book[3]
    }

    void testFindRowByCellValue() {
        def resultRow = sheet.findRowByCellValue('F4', 'Ken')
        assert resultRow.label == "7"
        assert resultRow.F_.value == "Ken"
    }

    void testFindRowByCellValue_prefixMatch() {
        def resultRow = sheet.findRowByCellValue('C4', 'Add')
        assert resultRow.label == "4"
        assert resultRow.C_.value == "Add Feature"
    }

    void testFindRowByCellValue_suffixMatch() {
        def resultRow = sheet.findRowByCellValue('H6', 'ed')
        assert resultRow.label == "6"
        assert resultRow.H_.value == "Closed"
    }

    void testFindRowByCellValue_numberMatch() {
        def resultRow = sheet.findRowByCellValue('I4', 20)
        assert resultRow.label == "5"
        assert resultRow.I_.value == 20
    }

    void testFindRowByEmptyCell_formulaValueColumn() {
        def emptyRow = sheet.findRowByEmptyCell('A2')
        assert emptyRow.label == "11"
    }

    void testFindRowByEmptyCell_valueColumn() {
        def emptyRow = sheet.findRowByEmptyCell('B3')
        assert emptyRow.label == "9"
    }

    void testFindAllRowsByCellValue() {
        def resultRows = sheet.findAllRowsByCellValue('D4', 'B')
        assert resultRows*.label == ["5", "6"]
        assert resultRows.every { it.D_.value == "B" }
    }

    void testFindAllRowsByCellValue_prefixMatch() {
        def resultRows = sheet.findAllRowsByCellValue('C4', 'Spec')
        assert resultRows*.label == ["5", "7"]
        assert resultRows.every { it.C_.value == "Spec Change" }
    }

    void testFindAllRowsByCellValue_numberMatch() {
        def resultRows = sheet.findAllRowsByCellValue('I4', 10)
        assert resultRows*.label == ["7", "8"]
        assert resultRows.every { it.I_.value == 10 }
    }
}

