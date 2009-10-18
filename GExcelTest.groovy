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
}
