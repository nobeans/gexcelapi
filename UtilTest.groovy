class UtilTest extends GroovyTestCase {

    void testConvertColLabelToNumber() throws Exception {
          assert Util.convertColLabelToNumber('A') == 0
          assert Util.convertColLabelToNumber('B') == 1
          assert Util.convertColLabelToNumber('C') == 2
          assert Util.convertColLabelToNumber('Z') == 25
          assert Util.convertColLabelToNumber('AA') == 26
          assert Util.convertColLabelToNumber('AB') == 27
          assert Util.convertColLabelToNumber('BA') == 52
          assert Util.convertColLabelToNumber('BZ') == 77
          assert Util.convertColLabelToNumber('ZZ') == 701
          assert Util.convertColLabelToNumber('AAA') == 702
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
