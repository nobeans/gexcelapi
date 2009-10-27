class GExcelTest extends GroovyTestCase {

    def book, sheet
    void setUp() throws Exception {
        book = GExcel.load("sample.xls")
        sheet = book[0]
    }

    void testLoad() throws Exception {
        assert GExcel.load("sample.xls")
        assert GExcel.load(new File("sample.xls"))
        new File("sample.xls").withInputStream { is ->
            assert GExcel.load(is)
        }
    }

    void testAccessSheet() throws Exception {
        assert sheet != book[1]

        def sheet3 = book.Sheet3 // by sheet name
        assert sheet3 == book[2]
        assert sheet3 == book["Sheet3"]
    }

    void testAccessCellValue() throws Exception {
        assert sheet["A1"].value == "Sheet1-A1"
 
        assert sheet.A1.value == "Sheet1-A1"
        sheet.A1.value = "MODIFIED"
        assert sheet.A1.value == "MODIFIED"
        sheet.A1 << "MORE-MODIFIED"
        assert sheet.A1.value == "MORE-MODIFIED"

        assert sheet.B1.value == "B1の内容"
        assert sheet.B2.value == "B2の内容"
    }

    void testAccessCellsOfSomeTypes() throws Exception {
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

    void testAccessCellsOfSomeTypes_asType() throws Exception {
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

    void testToString() throws Exception {
        assert sheet.A1.toString() == "Sheet1-A1"
        assert sheet.A2.toString() == "あいうえお"
        assert sheet.A3.toString() == "1234.0"
        assert sheet.A4.toString() == "40108.0" // Date is not assigned any cell type
        assert sheet.A5.toString() == "TRUE"
        assert sheet.A6.toString() == "FALSE"
    }

    void testRows() throws Exception {
        def rows = sheet.rows
        assert (rows[0].getCell(0) as String) == "Sheet1-A1"
        assert (rows[1].getCell(0) as String)  == "あいうえお"
        assert (rows[2].getCell(0) as Double) == 1234.0
        assert (rows[3].getCell(0) as Date) as String == "Thu Oct 22 00:00:00 JST 2009"
        assert (rows[4].getCell(0) as Boolean) == true
        assert (rows[5].getCell(0) as Boolean) == false
        assert  rows[6].getCell(0).cellFormula == "1+1"
        assert rows.size() == 7
        assert sheet.rows == sheet.rows() // both like property access and like method call
    }

//    void testRowWildcard() throws Exception {
//        def row1 = sheet.rows[0]
//        assert sheet._1 == row1
//
//        assert row1.A_.value == "Sheet1-A1" // A1
//        assert row1.B_.value == "B1の内容" // B1
//    }

//    void testIterateOfRows() throws Exception {
//        sheet.rows { row -> row.A_.value != "" }.each { row ->
//            println row.A_.value + ":" + row.B_.value
//        }
//    }

//    void testCols() throws Exception {
//        def colA = sheet.cols[0]
//        assert colA[0] == "" // A1
//        assert colA[1] == "" // A2
//    }
//
//    void testWildcard() throws Exception {
//        def row1 = sheet.rows[0]
//        assert sheet._1 == row1
//
//        def colA = sheet.cols[0]
//        assert sheet.A_  == colA
//    }
//
//    void testIterateOfCols() throws Exception {
//        def row1 = sheet.rows[0]
////        def row1 = sheet._1
////        def row1 = sheet.A_
//
//        row1.cells { cell -> cell.value != "" }.each { cell ->
//            println cell.value + ":" + cell.value
//        }
//    }

    //void testRange() throws Exception {
    //    def book = GExcel.load("sample.xls")
    //    def sheet = book[0]
    //    assert sheet[A1..B5]
    //    sheet[A1..B5].filledBy 0
    //    assert sheet[A1..B5]
    //}
}
