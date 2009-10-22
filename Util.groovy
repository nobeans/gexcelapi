class Util {

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

    private static convertNumberToColLabel(col) {
        def result = []
        int q = 1 // mandatory first loop
        int dividend = col
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
    static colIndex(cellLabel) {
        def matcher = (cellLabel =~ /([a-zA-Z]+)([0-9]+)/)
        convertColLabelToNumber(matcher[0][1])
    }

    // "A8" -> 8 -> 7
    static rowIndex(cellLabel) {
        def matcher = (cellLabel =~ /([a-zA-Z]+)([0-9]+)/)
        (matcher[0][2] as int) - 1
    }

    // (7, 0) -> "A8"
    static cellLabel(row, col) {
        convertNumberToColLabel(col) + (row + 1)
    }
}
