class UtilTest extends GroovyTestCase {

    void testConvertRowLabelToNumber() throws Exception {
          assert Util.convertRowLabelToNumber('A') == 0
          assert Util.convertRowLabelToNumber('B') == 1
          assert Util.convertRowLabelToNumber('C') == 2
          assert Util.convertRowLabelToNumber('Z') == 25
          assert Util.convertRowLabelToNumber('AA') == 26
          assert Util.convertRowLabelToNumber('AB') == 27
          assert Util.convertRowLabelToNumber('BA') == 52
          assert Util.convertRowLabelToNumber('BZ') == 77
          assert Util.convertRowLabelToNumber('ZZ') == 701
          assert Util.convertRowLabelToNumber('AAA') == 702
    }

    void testParseRowIndex() throws Exception {
        assert Util.rowIndex("A1") == 0
        assert Util.rowIndex("D2") == 1
        assert Util.rowIndex("Z1234") == 1233
    }

    void testParseColIndex() throws Exception {
        assert Util.colIndex("A2") == 0
        assert Util.colIndex("B2") == 1
        assert Util.colIndex("Z2") == 25
        assert Util.colIndex("AA88") == 26
        assert Util.colIndex("ZZ88") == 701
        assert Util.colIndex("AAA88") == 702
    }

}
