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

    def sampleFile = "build/classes/test/sample.xls"
    def book
    def sheet

    void setUp() {
        book = GExcel.open(sampleFile)
        sheet = book[0]
    }

    void tearDown() {
        book = null
        sheet = null
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
        assert sheet.A4.dateCellValue as String == "Thu Oct 22 00:00:00 JST 2009"
        assert sheet.A4.dateCellValue in java.util.Date

        assert sheet.A5.value == true
        assert sheet.A6.value == false

        try {
            sheet.A7.value
            fail()
        } catch (Exception e) {
            assert e.message == "unsupported cell type: 2"
        }
    }

    void testAccessCellsOfSomeTypes_asType() {
        assert (sheet.A1 as String) == "Sheet1-A1"
        assert (sheet.A1 as String) in String

        assert (sheet.A2 as String)  == "あいうえお"
        assert (sheet.A2 as String) in String

        assert (sheet.A3 as Double) == 1234.0
        assert (sheet.A3 as Double) in Double
        assert (sheet.A3 as Integer) == 1234
        assert (sheet.A3 as Integer) in Integer

        assert (sheet.A4 as Date) as String == "Thu Oct 22 00:00:00 JST 2009"
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

    void testWildcardOfRowAndCell() {
        assert sheet._1 == sheet.getRow(0) // not CellRange
        assert sheet._1.A_.value == "Sheet1-A1" // A1
        assert sheet._1.B_.value == "B1の内容" // B1

        assert sheet._2 == sheet.getRow(1)
        assert sheet._2.A_.value == "あいうえお" // A2
    }

    void testWildcardOfColumn() {
        assert sheet.A_ in CellRange
        assert sheet.A_ == sheet.A1_A7
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
        assert sheet._1.validate() // row
        assert sheet.validate()    // sheet

        sheet.B1.validators << { false } // force invalid
        assert sheet.B1.validate() == false
        assert sheet._1.validate() == false // recursive
        assert sheet.validate() == false    // recursive
    }

    void testCellRange() {
        assert sheet.A1_B2.collect { it?.value } == ["Sheet1-A1", "あいうえお", "B1の内容", "B2の内容"]
        assert sheet.A1_B3.collect { it?.value } == ["Sheet1-A1", "あいうえお", 1234.0, "B1の内容", "B2の内容", null]

        assert sheet.A2_B3.each { it?.value = "X" }
        assert sheet.A1_B3.collect { it?.value } == ["Sheet1-A1", "X", "X", "B1の内容", "X", null] // cannot apply a value to a null cell
    }

    void testLabel_forCell() {
        println sheet.A1
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
}

