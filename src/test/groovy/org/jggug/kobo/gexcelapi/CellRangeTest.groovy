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

class CellRangeTest extends GroovyTestCase {

    def sampleFile = "build/classes/test/sample.xls"
    def book, sheet

    void setUp() {
        book = GExcel.open(sampleFile)
        sheet = book[0]
    }

    void tearDown() {
        book = null
        sheet = null
    }

    void testRectangle_A1_B3() {
        def expected = [
            [sheet.A1, sheet.B1],
            [sheet.A2, sheet.B2],
            [sheet.A3, sheet.B3]
        ]
        assert CellRange.newRectangleCellRange(sheet, "A1", "B3") == expected
        assert new CellRange(sheet, "A1", "B3") == expected
        assert new CellRange(sheet, 0, 0, 2, 1) == expected
        assert new CellRange(sheet, "A1", "B3", false) == expected
        assert new CellRange(sheet, 0, 0, 2, 1, false) == expected
    }

    void testRectangle_A1_A3() {
        def expected = [
            [sheet.A1],
            [sheet.A2],
            [sheet.A3]
        ]
        assert CellRange.newRectangleCellRange(sheet, "A1", "A3") == expected
        assert new CellRange(sheet, "A1", "A3") == expected
        assert new CellRange(sheet, 0, 0, 2, 0) == expected
        assert new CellRange(sheet, "A1", "A3", false) == expected
        assert new CellRange(sheet, 0, 0, 2, 0, false) == expected
    }

    void testRectangle_A1_C1() {
        def expected = [
            [sheet.A1, sheet.B1, sheet.C1]
        ]
        assert CellRange.newRectangleCellRange(sheet, "A1", "C1") == expected
        assert new CellRange(sheet, "A1", "C1") == expected
        assert new CellRange(sheet, 0, 0, 0, 2) == expected
        assert new CellRange(sheet, "A1", "C1", false) == expected
        assert new CellRange(sheet, 0, 0, 0, 2, false) == expected
    }

    void testSequential_A1_B3() {
        def expected = [
            sheet.A1, sheet.B1,
            sheet.A2, sheet.B2,
            sheet.A3, sheet.B3
        ]
        assert CellRange.newSequentialCellRange(sheet, "A1", "B3") == expected
        assert new CellRange(sheet, "A1", "B3", true) == expected
        assert new CellRange(sheet, 0, 0, 2, 1, true) == expected
    }

    void testSequential_A1_A3() {
        def expected = [
            sheet.A1,
            sheet.A2,
            sheet.A3
        ]
        assert CellRange.newSequentialCellRange(sheet, "A1", "A3") == expected
        assert new CellRange(sheet, "A1", "A3", true) == expected
        assert new CellRange(sheet, 0, 0, 2, 0, true) == expected
    }

    void testSequential_A1_C1() {
        def expected = [
            sheet.A1, sheet.B1, sheet.C1
        ]
        assert CellRange.newSequentialCellRange(sheet, "A1", "C1") == expected
        assert new CellRange(sheet, "A1", "C1", true) == expected
        assert new CellRange(sheet, 0, 0, 0, 2, true) == expected
    }

    void testLabel() {
        assert sheet.A1_B3.label == "A1:B3"
        assert sheet.A1_A3.label == "A1:A3"
        assert sheet.A1_C1.label == "A1:C1"
    }

    void testToHtml_A1_C1_withTitle() {
        sheet = book[1] // having merged regions
        def html = sheet.A1_C6.toHtml("test")
        assert html =~ "<html>"
        assert html =~ "<meta http-equiv='Content-Type' content='text/html; charset=UTF-8' />"
        assert html =~ "<title>test</title>"
        // snip checking contents
    }

    void testToHtml_A1_C1_withoutTitle() {
        sheet = book[1] // having merged regions
        def html = sheet.A1_C6.toHtml()
        assert html =~ "<html>"
        assert html =~ "<meta http-equiv='Content-Type' content='text/html; charset=UTF-8' />"
        assert html =~ "<title>A1:C6 from Excel</title>"
        // snip checking contents
    }

    void testToHtml_A1_C1_wichCharset() {
        sheet = book[1] // having merged regions
        def html = sheet.A1_C6.toHtml("test", "iso-8859-1")
        assert html =~ "<html>"
        assert html =~ "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1' />"
        assert html =~ "<title>test</title>"
        // snip checking contents
    }
}
