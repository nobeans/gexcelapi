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

class RectangleCellRangeTest extends GroovyTestCase {

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
        assert new RectangleCellRange(sheet, "A1", "B3") == [
            [sheet.A1, sheet.B1],
            [sheet.A2, sheet.B2],
            [sheet.A3, sheet.B3]
        ]
        assert new RectangleCellRange(sheet, "A1", "B2") == [
            [sheet.A1, sheet.B1],
            [sheet.A2, sheet.B2]
        ]
    }

    void testIterator_points() {
        assert new RectangleCellRange(sheet, 0, 0, 2, 1) == [
            [sheet.A1, sheet.B1],
            [sheet.A2, sheet.B2],
            [sheet.A3, sheet.B3]
        ]
        assert new RectangleCellRange(sheet, 0, 0, 1, 1) == [
            [sheet.A1, sheet.B1],
            [sheet.A2, sheet.B2]
        ]
    }

    void testForIn_A1_B2() {
        def range = new RectangleCellRange(sheet, "A1", "B2") // A1..B2
        def result = []
        for (cell in range) {
            result << cell
        }
        assert result == [[sheet.A1, sheet.B1], [sheet.A2, sheet.B2]]
    }

}
