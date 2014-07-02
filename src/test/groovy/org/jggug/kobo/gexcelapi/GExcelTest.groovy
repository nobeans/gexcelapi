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

class GExcelTest extends GroovyTestCase {

    def sampleFile = "build/resources/test/sample.xls"
    def outputPath = "output.xls"
    def outputFile
    def book
    def sheet
    def sheet5

    void setUp() {
        book = GExcel.open(sampleFile)
        sheet = book[0]
        sheet5 = book["Sheet5"]
        outputFile = new File(outputPath)

        // dateCellValue is affected by TimeZone.
        // So it should be explicitly set as GMT in order to avoid causing a failure dependent on an environment.
        TimeZone.default = TimeZone.getTimeZone("GMT")
    }

    void tearDown() {
        book = null
        sheet = null
        sheet5 = null
        outputFile.delete()
    }

    void testOpen() {
        assert GExcel.open(sampleFile)
        assert GExcel.open(new File(sampleFile))
        new File(sampleFile).withInputStream { is ->
            assert GExcel.open(is)
        }
    }

    void testAccessSheet() {
        assert sheet != book[1]

        def sheet3 = book.Sheet3 // by sheet name
        assert sheet3 == book[2]
        assert sheet3 == book["Sheet3"]
    }

    void testAccessCellValue() {
        assert sheet["A1"].value == "Sheet1-A1"

        assert sheet.A1.value == "Sheet1-A1"
        sheet.A1.value = "MODIFIED"
        assert sheet.A1.value == "MODIFIED"
        sheet.A1 << "MORE-MODIFIED"
        assert sheet.A1.value == "MORE-MODIFIED"

        assert sheet.B1.value == "B1の内容"
        assert sheet.B2.value == "B2の内容"
    }

    void testAccessCellsOfSomeTypes() {
        assert sheet.A1.value == "Sheet1-A1"
        assert sheet.A1.value in String

        assert sheet.A2.value == "あいうえお"
        assert sheet.A2.value in String

        assert sheet.A3.value == 1234.0
        assert sheet.A3.value in Double
        assert sheet.A3.value.intValue() in Integer // method to get int value

        // it cannot implicitly accessing value, so explicitly access.
        assert sheet.A4.dateCellValue.format("yyyy-MM-dd HH:mm:ss", TimeZone.getTimeZone("GMT")) == "2009-10-22 00:00:00"
        assert sheet.A4.dateCellValue.format("yyyy-MM-dd HH:mm:ss", TimeZone.getTimeZone("JST")) == "2009-10-22 09:00:00"
        assert sheet.A4.dateCellValue in java.util.Date

        assert sheet.A5.value == true
        assert sheet.A6.value == false

        // if you want to get result of formula cell, you can use the dedicated accessor
        // method of POI, like getNumericCellValue().
        assert sheet.A7.numericCellValue == 2
        assert sheet.A7.value == "1+1" // if you wan to get the formula, you can use 'value'.
    }

    void testAccessCellsOfSomeTypes_asType() {
        assert (sheet.A1 as String) == "Sheet1-A1"
        assert (sheet.A1 as String) in String

        assert (sheet.A2 as String) == "あいうえお"
        assert (sheet.A2 as String) in String

        assert (sheet.A3 as Double) == 1234.0
        assert (sheet.A3 as Double) in Double
        assert (sheet.A3 as Integer) == 1234
        assert (sheet.A3 as Integer) in Integer

        assert (sheet.A4 as Date).format("yyyy-MM-dd HH:mm:ss", TimeZone.getTimeZone("GMT")) == "2009-10-22 00:00:00"
        assert (sheet.A4 as Date).format("yyyy-MM-dd HH:mm:ss", TimeZone.getTimeZone("JST")) == "2009-10-22 09:00:00"
        assert (sheet.A4 as Date) in java.util.Date

        assert (sheet.A5 as Boolean) == true
        assert (sheet.A6 as Boolean) == false

        try {
            sheet.A1 as List
            fail()
        } catch (Exception e) {
            assert e.message == "unsupported cell type: 1"
        }
    }

    void testInspectCellType() {
        assert sheet.A1.isStringType()
        assert sheet.A2.isStringType()
        assert sheet.A3.isNumericType()
        assert sheet.A4.isNumericType() // wrong: not numeric but date
        assert sheet.A5.isBooleanType()
        assert sheet.A6.isBooleanType()
    }

    void testToString() {
        assert sheet.A1.toString() == "Sheet1-A1"
        assert sheet.A2.toString() == "あいうえお"
        assert sheet.A3.toString() == "1234.0"
        assert sheet.A4.toString() == "40108.0" // Date is not assigned any cell type
        assert sheet.A5.toString() == "TRUE"
        assert sheet.A6.toString() == "FALSE"
        assert sheet.A7.toString() == "1+1"
    }

    void testRows() {
        def rows = sheet.rows
        assert rows[0] == sheet.getRow(0)
        assert rows[1] == sheet.getRow(1)
        assert rows[2] == sheet.getRow(2)
        assert rows[3] == sheet.getRow(3)
        assert rows[4] == sheet.getRow(4)
        assert rows[5] == sheet.getRow(5)
        assert rows[6] == sheet.getRow(6)
        assert rows.size() == 7

        assert sheet.rows == sheet.rows() // both like property access and like method call
    }

    void testCellRange_wildcardOfRow() {
        assert sheet._1 == [sheet.A1, sheet.B1]
        assert sheet._1 in CellRange

        assert sheet._2 == [sheet.A2, sheet.B2]
        assert sheet._2 in CellRange
    }

    void testCellRange_wildcardOfColumn() {
        assert sheet.A_ in CellRange
        assert sheet.A_ == [sheet.A1, sheet.A2, sheet.A3, sheet.A4, sheet.A5, sheet.A6, sheet.A7]
    }

    void testRectangleCellRange() {
        assert sheet.A1_B2 == [[sheet.A1, sheet.B1], [sheet.A2, sheet.B2]]
        assert sheet.A1_B3 == [[sheet.A1, sheet.B1], [sheet.A2, sheet.B2], [sheet.A3, sheet.B3]]
        assert sheet.A1_A3 == [[sheet.A1], [sheet.A2], [sheet.A3]]
        assert sheet.A1_C1 == [[sheet.A1, sheet.B1, sheet.C1]]
    }

    void testHowToUseRowIterator() {
        assert (
            sheet.findAll { row -> row.A_.isStringType() }.collect { row -> row.A_.value + "," + row.B_.value }
        ) == ["Sheet1-A1,B1の内容", "あいうえお,B2の内容"]
    }

    void testValidationOfCell() {
        sheet.A1.validators << { it.value == "Sheet1-A1" }
        assert sheet.A1.validate()
        assert sheet.A1.validators.size() == 1

        sheet.A1.addValidator { it.value in String } // add
        assert sheet.A1.validate()
        assert sheet.A1.validators.size() == 2

        sheet.A1.validators << { it.value in Integer } // wrong
        assert sheet.A1.validate() == false
        assert sheet.A1.validators.size() == 3

        sheet.A1.validator = { it.value == "Sheet1-A1" }
        assert sheet.A1.validate()
        assert sheet.A1.validators.size() == 1

        sheet.A1.setValidator { it.value in String } // setter
        assert sheet.A1.validate()
        assert sheet.A1.validators.size() == 1
    }

    void testValidationOfRowAndSheet() {
        sheet.A1.validators << { it.value == "Sheet1-A1" }
        sheet.B1.validators << { it.value == "B1の内容" }
        assert sheet.A1.validate()
        assert sheet.B1.validate()
        assert sheet.getRow(0).validate() // row
        assert sheet._1.validate() // sequential CellRange for row
        assert sheet.A_.validate() // sequential CellRange for column
        assert sheet.A1_B2.validate() // rectangle CellRange for column
        assert sheet.validate()    // sheet

        sheet.B1.validators << { false } // force invalid
        assert sheet.B1.validate() == false
        assert sheet._1.validate() == false // recursive
        assert sheet.validate() == false    // recursive
    }

    void testLabel_forCell() {
        assert sheet.A1.label == "A1"
        assert sheet.A2.label == "A2"
        assert sheet.A3.label == "A3"
        assert sheet.B1.label == "B1"
        assert sheet.B2.label == "B2"
    }

    void testColumnLabel_forCell() {
        assert sheet.A1.columnLabel == "A"
        assert sheet.A2.columnLabel == "A"
        assert sheet.A3.columnLabel == "A"
        assert sheet.B1.columnLabel == "B"
        assert sheet.B2.columnLabel == "B"
    }

    void testRowLabel_forCell() {
        assert sheet.A1.rowLabel == "1"
        assert sheet.A2.rowLabel == "2"
        assert sheet.A3.rowLabel == "3"
        assert sheet.B1.rowLabel == "1"
        assert sheet.B2.rowLabel == "2"
    }

    void testLabel_forRow() {
        assert sheet.getRow(0).label == "1"
        assert sheet.getRow(1).label == "2"
        assert sheet.getRow(2).label == "3"
    }

    void testMergedRegion_getEnclosingMergedRegion() {
        sheet = book[1] // having merged regions
        assert sheet.getEnclosingMergedRegion(sheet.A1) == null
        assert sheet.getEnclosingMergedRegion(sheet.B1) == null
        assert sheet.getEnclosingMergedRegion(sheet.C1) == null

        assert sheet.getEnclosingMergedRegion(sheet.A2).formatAsString() == "A2:C2"
        assert sheet.getEnclosingMergedRegion(sheet.B2).formatAsString() == "A2:C2"
        assert sheet.getEnclosingMergedRegion(sheet.C2).formatAsString() == "A2:C2"

        assert sheet.getEnclosingMergedRegion(sheet.C3).formatAsString() == "C3:C5"
        assert sheet.getEnclosingMergedRegion(sheet.C4).formatAsString() == "C3:C5"
        assert sheet.getEnclosingMergedRegion(sheet.C5).formatAsString() == "C3:C5"

        assert sheet.getEnclosingMergedRegion(sheet.A4).formatAsString() == "A4:B5"
        assert sheet.getEnclosingMergedRegion(sheet.B4).formatAsString() == "A4:B5"
        assert sheet.getEnclosingMergedRegion(sheet.A5).formatAsString() == "A4:B5"
        assert sheet.getEnclosingMergedRegion(sheet.B5).formatAsString() == "A4:B5"
    }

    void testMergedRegion_width_height() {
        sheet = book[1] // having merged regions
        assert sheet.getEnclosingMergedRegion(sheet.A2).height == 1
        assert sheet.getEnclosingMergedRegion(sheet.A2).width == 3

        assert sheet.getEnclosingMergedRegion(sheet.C3).height == 3
        assert sheet.getEnclosingMergedRegion(sheet.C3).width == 1

        assert sheet.getEnclosingMergedRegion(sheet.A4).height == 2
        assert sheet.getEnclosingMergedRegion(sheet.A4).width == 2
    }

    void testMergedRegion_isFirstCell() {
        sheet = book[1] // having merged regions
        assert sheet.getEnclosingMergedRegion(sheet.A2).isFirstCell(sheet.A2)
        assert sheet.getEnclosingMergedRegion(sheet.A2).isFirstCell(sheet.B2) == false
        assert sheet.getEnclosingMergedRegion(sheet.A2).isFirstCell(sheet.C2) == false
        assert sheet.getEnclosingMergedRegion(sheet.A2).isFirstCell(sheet.D2) == false // out of the range

        assert sheet.getEnclosingMergedRegion(sheet.C3).isFirstCell(sheet.C3)
        assert sheet.getEnclosingMergedRegion(sheet.C3).isFirstCell(sheet.C4) == false
        assert sheet.getEnclosingMergedRegion(sheet.C3).isFirstCell(sheet.C5) == false
        assert sheet.getEnclosingMergedRegion(sheet.C3).isFirstCell(sheet.C6) == false // out of the range

        assert sheet.getEnclosingMergedRegion(sheet.A4).isFirstCell(sheet.A4)
        assert sheet.getEnclosingMergedRegion(sheet.A4).isFirstCell(sheet.B4) == false
        assert sheet.getEnclosingMergedRegion(sheet.A4).isFirstCell(sheet.A5) == false
        assert sheet.getEnclosingMergedRegion(sheet.A4).isFirstCell(sheet.B5) == false
    }

    void testFindByCellValue() {
        def resultRow = sheet5.findByCellValue('F4', 'Ken');
        assert resultRow != null
        assert resultRow.label == "7"  // label equals line number on excel
        assert resultRow.F_.value == "Ken"
        assert resultRow.getCell(5) == sheet5.F7
    }

    void testFindByCellValuePrefixMatch() {
        def resultRow = sheet5.findByCellValue('C4', 'Add');
        assert resultRow != null
        assert resultRow.label == "4"  // label equals line number on excel
        assert resultRow.C_.value == "Add Feature"
        assert resultRow.getCell(2) == sheet5.C4
    }

    void testFindByCellValueSuffixMatch() {
        def resultRow = sheet5.findByCellValue('H6', 'ed');
        assert resultRow != null
        assert resultRow.label == "6"  // label equals line number on excel
        assert resultRow.H_.value == "Closed"
        assert resultRow.getCell(7) == sheet5.H6
    }

    void testFindByCellValueNumberMatch() {
        def resultRow = sheet5.findByCellValue('I4', 20);
        assert resultRow != null
        assert resultRow.label == "5"  // label equals line number on excel
        assert resultRow.I_.value == 20
        assert resultRow.getCell(8) == sheet5.I5
    }

    void testFindAllByCellValue() {
        def resultRows = sheet5.findAllByCellValue('D4', 'B');
        assert resultRows[0].label == "5"
        assert resultRows[0].D_.value == "B"
        assert resultRows[1].label == "6"
        assert resultRows[1].D_.value == "B"
    }

    void testFindAllByCellValuePrefixMatch() {
        def resultRows = sheet5.findAllByCellValue('C4', 'Spec');
        assert resultRows[0].label == "5"
        assert resultRows[0].C_.value == "Spec Change"
        assert resultRows[1].label == "7"
        assert resultRows[1].C_.value == "Spec Change"
    }

    void testFindAllByCellValueNumberMatch() {
        def resultRows = sheet5.findAllByCellValue('I4', 10);
        assert resultRows[0].label == "7"
        assert resultRows[0].I_.value == 10
        assert resultRows[1].label == "8"
        assert resultRows[1].I_.value == 10
    }
}

