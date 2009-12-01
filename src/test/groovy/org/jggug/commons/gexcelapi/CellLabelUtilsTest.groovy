package org.jggug.commons.gexcelapi

class CellLabelUtilsTest extends GroovyTestCase {

    void testConvertColLabelToNumber() {
          assert CellLabelUtils.convertColLabelToNumber('A') == 0
          assert CellLabelUtils.convertColLabelToNumber('B') == 1
          assert CellLabelUtils.convertColLabelToNumber('C') == 2
          assert CellLabelUtils.convertColLabelToNumber('Z') == 25
          assert CellLabelUtils.convertColLabelToNumber('AA') == 26
          assert CellLabelUtils.convertColLabelToNumber('AB') == 27
          assert CellLabelUtils.convertColLabelToNumber('BA') == 52
          assert CellLabelUtils.convertColLabelToNumber('BZ') == 77
          assert CellLabelUtils.convertColLabelToNumber('ZZ') == 701
          assert CellLabelUtils.convertColLabelToNumber('AAA') == 702
    }

    void testParseRowIndex() {
        assert CellLabelUtils.rowIndex("A1") == 0
        assert CellLabelUtils.rowIndex("D2") == 1
        assert CellLabelUtils.rowIndex("Z1234") == 1233
        assert CellLabelUtils.rowIndex("_1234") == 1233 // wildcard
    }

    void testParseColIndex() {
        assert CellLabelUtils.colIndex("A2") == 0
        assert CellLabelUtils.colIndex("B2") == 1
        assert CellLabelUtils.colIndex("Z2") == 25
        assert CellLabelUtils.colIndex("AA88") == 26
        assert CellLabelUtils.colIndex("ZZ88") == 701
        assert CellLabelUtils.colIndex("AAA88") == 702
        assert CellLabelUtils.colIndex("AAA_") == 702 // wildcard
    }

    void testConvertNumberToColLabel() {
          assert CellLabelUtils.convertNumberToColLabel(0) == 'A'
          assert CellLabelUtils.convertNumberToColLabel(1) == 'B'
          assert CellLabelUtils.convertNumberToColLabel(2) == 'C'
          assert CellLabelUtils.convertNumberToColLabel(25) == 'Z'
          assert CellLabelUtils.convertNumberToColLabel(26) == 'AA'
          assert CellLabelUtils.convertNumberToColLabel(27) == 'AB'
          assert CellLabelUtils.convertNumberToColLabel(52) == 'BA'
          assert CellLabelUtils.convertNumberToColLabel(77) == 'BZ'
          assert CellLabelUtils.convertNumberToColLabel(701) == 'ZZ'
          assert CellLabelUtils.convertNumberToColLabel(702) == 'AAA'
    }

    void testCellLabel() {
        assert CellLabelUtils.cellLabel(2, 0) == "A3"
        assert CellLabelUtils.cellLabel(3, 1) == "B4"
        assert CellLabelUtils.cellLabel(4, 25) == "Z5"
        assert CellLabelUtils.cellLabel(80, 26) == "AA81"
        assert CellLabelUtils.cellLabel(81, 701) == "ZZ82"
        assert CellLabelUtils.cellLabel(82, 702) == "AAA83"
    }

}
