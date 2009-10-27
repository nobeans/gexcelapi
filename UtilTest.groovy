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
        assert Util.rowIndex("_1234") == 1233 // wildcard
    }

    void testParseColIndex() throws Exception {
        assert Util.colIndex("A2") == 0
        assert Util.colIndex("B2") == 1
        assert Util.colIndex("Z2") == 25
        assert Util.colIndex("AA88") == 26
        assert Util.colIndex("ZZ88") == 701
        assert Util.colIndex("AAA88") == 702
        assert Util.colIndex("AAA_") == 702 // wildcard
    }

    void testConvertNumberToColLabel() throws Exception {
          assert Util.convertNumberToColLabel(0) == 'A'
          assert Util.convertNumberToColLabel(1) == 'B'
          assert Util.convertNumberToColLabel(2) == 'C'
          assert Util.convertNumberToColLabel(25) == 'Z'
          assert Util.convertNumberToColLabel(26) == 'AA'
          assert Util.convertNumberToColLabel(27) == 'AB'
          assert Util.convertNumberToColLabel(52) == 'BA'
          assert Util.convertNumberToColLabel(77) == 'BZ'
          assert Util.convertNumberToColLabel(701) == 'ZZ'
          assert Util.convertNumberToColLabel(702) == 'AAA'
    }

    void testCellLabel() throws Exception {
        assert Util.cellLabel(2, 0) == "A3"
        assert Util.cellLabel(3, 1) == "B4"
        assert Util.cellLabel(4, 25) == "Z5"
        assert Util.cellLabel(80, 26) == "AA81"
        assert Util.cellLabel(81, 701) == "ZZ82"
        assert Util.cellLabel(82, 702) == "AAA83"
    }

}
