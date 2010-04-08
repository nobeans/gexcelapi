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

    void testIterator_label() {
        assert new CellRange(sheet, "A1", "B3").collect { it?.value } == ["Sheet1-A1", "あいうえお", 1234.0, "B1の内容", "B2の内容", null]
        assert new CellRange(sheet, "A1", "B2").collect { it?.value } == ["Sheet1-A1", "あいうえお", "B1の内容", "B2の内容"]
    }

    void testIterator_points() {
        assert new CellRange(sheet, 0, 0, 2, 1).collect { it?.value } == ["Sheet1-A1", "あいうえお", 1234.0, "B1の内容", "B2の内容", null]
        assert new CellRange(sheet, 0, 0, 1, 1).collect { it?.value } == ["Sheet1-A1", "あいうえお", "B1の内容", "B2の内容"]
    }

    void testnScope_A1_B2() {
        def range = new CellRange(sheet, "A1", "B2") // A1..B2
        assert sheet.A1 in range
        assert sheet.A2 in range
        assert sheet.B1 in range
        assert sheet.B2 in range
        assert sheet.C1 in range == false
    }

    void testForIn_A1_B2() {
        def range = new CellRange(sheet, "A1", "B2") // A1..B2
        def result = []
        for (cell in range) {
            result << cell?.value
        }
        assert result == ["Sheet1-A1", "あいうえお", "B1の内容", "B2の内容"]
    }

}
