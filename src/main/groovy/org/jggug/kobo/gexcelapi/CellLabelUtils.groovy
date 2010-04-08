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

class CellLabelUtils {

    private static final int ORIGIN = ('A' as char) as int
    private static final int RADIX = 26

    private static convertColLabelToNumber(ascii) {
        def num = 0
        ascii.toUpperCase().reverse().eachWithIndex { ch, i ->
            def delta = ((ch as char) as int) - ORIGIN + 1
            num += delta * (RADIX**i)
        }
        num - 1 // convert for "index" which starts from 0
    }

    private static convertNumberToColLabel(column) {
        def result = []
        int q = 1 // mandatory first loop
        int dividend = column
        while (q > 0) {
            q = dividend / RADIX
            int r = dividend % RADIX
            result << String.valueOf((r + ORIGIN) as char)
            //println "$dividend, $RADIX => $q ($r) : $result"
            dividend = q - 1 // 0 == A
        }
        return result.reverse().join()
    }

    // "A8" -> A -> 0
    static columnIndex(cellLabel) {
        def matcher = (cellLabel =~ /([a-zA-Z]+)(_|[0-9]+)/)
        convertColLabelToNumber(matcher[0][1])
    }

    // "A8" -> 8 -> 7
    static rowIndex(cellLabel) {
        def matcher = (cellLabel =~ /(_|[a-zA-Z]+)([0-9]+)/)
        (matcher[0][2] as int) - 1
    }

    // (7, 0) -> "A8"
    static cellLabel(row, column) {
        convertNumberToColLabel(column) + (row + 1)
    }
}
