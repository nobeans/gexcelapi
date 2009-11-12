package org.jggug.commons.gexcelapi

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
