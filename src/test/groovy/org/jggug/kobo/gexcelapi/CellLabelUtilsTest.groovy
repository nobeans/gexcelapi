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

import org.jggug.kobo.gexcelapi.CellLabelUtils as CLU

class CellLabelUtilsTest extends GroovyTestCase {

    void testConvertColLabelToNumber() {
          assert CLU.convertColLabelToNumber('A') == 0
          assert CLU.convertColLabelToNumber('B') == 1
          assert CLU.convertColLabelToNumber('C') == 2
          assert CLU.convertColLabelToNumber('Z') == 25
          assert CLU.convertColLabelToNumber('AA') == 26
          assert CLU.convertColLabelToNumber('AB') == 27
          assert CLU.convertColLabelToNumber('BA') == 52
          assert CLU.convertColLabelToNumber('BZ') == 77
          assert CLU.convertColLabelToNumber('ZZ') == 701
          assert CLU.convertColLabelToNumber('AAA') == 702
    }

    void testParseRowIndex() {
        assert CLU.rowIndex("A1") == 0
        assert CLU.rowIndex("D2") == 1
        assert CLU.rowIndex("Z1234") == 1233
        assert CLU.rowIndex("_1234") == 1233 // wildcard
    }

    void testParseColIndex() {
        assert CLU.columnIndex("A2") == 0
        assert CLU.columnIndex("B2") == 1
        assert CLU.columnIndex("Z2") == 25
        assert CLU.columnIndex("AA88") == 26
        assert CLU.columnIndex("ZZ88") == 701
        assert CLU.columnIndex("AAA88") == 702
        assert CLU.columnIndex("AAA_") == 702 // wildcard
    }

    void testConvertNumberToColLabel() {
          assert CLU.convertNumberToColLabel(0) == 'A'
          assert CLU.convertNumberToColLabel(1) == 'B'
          assert CLU.convertNumberToColLabel(2) == 'C'
          assert CLU.convertNumberToColLabel(25) == 'Z'
          assert CLU.convertNumberToColLabel(26) == 'AA'
          assert CLU.convertNumberToColLabel(27) == 'AB'
          assert CLU.convertNumberToColLabel(52) == 'BA'
          assert CLU.convertNumberToColLabel(77) == 'BZ'
          assert CLU.convertNumberToColLabel(701) == 'ZZ'
          assert CLU.convertNumberToColLabel(702) == 'AAA'
    }

    void testCellLabel() {
        assert CLU.cellLabel(2, 0) == "A3"
        assert CLU.cellLabel(3, 1) == "B4"
        assert CLU.cellLabel(4, 25) == "Z5"
        assert CLU.cellLabel(80, 26) == "AA81"
        assert CLU.cellLabel(81, 701) == "ZZ82"
        assert CLU.cellLabel(82, 702) == "AAA83"
    }

}
