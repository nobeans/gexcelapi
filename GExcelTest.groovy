class GExcelTest extends GroovyTestCase {

    void testConvertRowLabelToNumber() throws Exception {
          assert GExcel.convertRowLabelToNumber('A') == 0
          assert GExcel.convertRowLabelToNumber('B') == 1
          assert GExcel.convertRowLabelToNumber('C') == 2
          assert GExcel.convertRowLabelToNumber('Z') == 25
          assert GExcel.convertRowLabelToNumber('AA') == 26
          assert GExcel.convertRowLabelToNumber('AB') == 27
          assert GExcel.convertRowLabelToNumber('BA') == 52
          assert GExcel.convertRowLabelToNumber('BZ') == 77
          assert GExcel.convertRowLabelToNumber('ZZ') == 701
          assert GExcel.convertRowLabelToNumber('AAA') == 702
    }

    void testParseRowIndex() throws Exception {
        assert GExcel.rowIndex("A1") == 0
        assert GExcel.rowIndex("D2") == 1
        assert GExcel.rowIndex("Z1234") == 1233
    }

    void testParseColIndex() throws Exception {
        assert GExcel.colIndex("A2") == 0
        assert GExcel.colIndex("B2") == 1
        assert GExcel.colIndex("Z2") == 25
        assert GExcel.colIndex("AA88") == 26
        assert GExcel.colIndex("ZZ88") == 701
        assert GExcel.colIndex("AAA88") == 702
    }

    void testLoad() throws Exception {
        assert GExcel.load("./sample.xls")
        assert GExcel.load(new File("sample.xls"))
        new File("sample.xls").withInputStream { is ->
            assert GExcel.load(is)
        }
    }

    void testAccessSheet() throws Exception {
        def book = GExcel.load(new File("sample.xls"))
        def sheet = book[0]
        assert sheet != book[1]

        def sheet3 = book.Sheet3 // by sheet name
        assert sheet3 == book[2]
        assert sheet3 == book["Sheet3"]
    }

    void testAccessCellValue() throws Exception {
        def book = GExcel.load(new File("sample.xls"))
        def sheet = book[0]
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
        def book = GExcel.load(new File("sample.xls"))
        def sheet = book[0]

        assert sheet.A1.value == "Sheet1-A1"
        assert sheet.A1.value in String

        assert sheet.A2.value == "あいうえお"
        assert sheet.A2.value in String

        assert sheet.A3.value == 1234.0
        assert sheet.A3.value in Double
        assert sheet.A3.value.intValue() in Integer // method to get int value

        // it cannot implicitly accessing value, so explicitly access.
        assert sheet.A4.dateCellValue in java.util.Date
        assert sheet.A4.dateCellValue as String == "Thu Oct 22 00:00:00 JST 2009"

        assert sheet.A5.value == true
        assert sheet.A6.value == false
    }

    void testAccessCellsOfSomeTypes_asType() throws Exception {
        def book = GExcel.load(new File("sample.xls"))
        def sheet = book[0]

        assert (sheet.A1 as String) == "Sheet1-A1"
        assert (sheet.A1 as String) in String

        assert (sheet.A2 as String)  == "あいうえお"
        assert (sheet.A2 as String) in String

        assert (sheet.A3 as Double) == 1234.0
        assert (sheet.A3 as Double) in Double
        assert (sheet.A3 as Integer) == 1234
        assert (sheet.A3 as Integer) in Integer

        assert (sheet.A4 as Date) in java.util.Date
        assert (sheet.A4 as Date) as String == "Thu Oct 22 00:00:00 JST 2009"

        assert (sheet.A5 as Boolean) == true
        assert (sheet.A6 as Boolean) == false
    }

    //void testRange() throws Exception {
    //    def book = GExcel.load(new File("sample.xls"))
    //    def sheet = book[0]
    //    assert sheet[A1..B5]
    //    sheet[A1..B5].filledBy 0
    //    assert sheet[A1..B5]
    //}
}
