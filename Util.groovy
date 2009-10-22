class Util {

    private static convertColLabelToNumber(ascii) {
        final origin = ('A' as char) as int
        final radix = 26
        def num = 0
        ascii.toUpperCase().reverse().eachWithIndex { ch, i ->
            def delta = ((ch as char) as int) - origin + 1
            num += delta * (radix**i)
        }
        num - 1 // convert for "index" which starts from 0
    }

    // "A8" -> A -> 0
    static colIndex(expr) {
        def matcher = (expr =~ /([a-zA-Z]+)([0-9]+)/)
        convertColLabelToNumber(matcher[0][1])
    }

    // "A8" -> 8 -> 7
    static rowIndex(expr) {
        def matcher = (expr =~ /([a-zA-Z]+)([0-9]+)/)
        (matcher[0][2] as int) - 1
    }
}
