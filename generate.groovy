// usage: groovy generate.groovy entity-def.xls

// ----------------------------
// Grapeによるライブラリ取得
@Grab(group='org.apache.poi', module='poi', version='3.5-beta3')
import org.apache.poi.hssf.usermodel.*
import org.apache.poi.poifs.filesystem.*
//groovy.grape.Grape.grab(group:'org.apache.poi', module:'poi', version:'3.0.2-FINAL')
//@Grab('org.apache.poi:poi:3.0.2-FINAL') // only for Groovy1.7

// ----------------------------
// エンティティ情報保持用のクラス
class Table {
    String physicalName = ""
    String logicalName = ""
    List<Column> columns = []
    public String toString() {
        "$logicalName($physicalName) {\n" +
        columns.collect{ "    ${it}" }.join("\n") +
        "\n}"
    }
}
class Column {
    String physicalName = ""
    String logicalName = ""
    String type = ""
    int length = 0
    boolean notNull = true
    public String toString() {
        "$logicalName($physicalName) [$type(${length})] ${notNull?'NOT NULL':''}"
    }
}

// ----------------------------
// Excelファイルの読み込み
def inputFile = new File(args[0])
def book = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(inputFile)))
def sheet = book.getSheetAt(0) // 今回は1番目のシートのみ対象とする
def cell = { row, col ->
    sheet.getRow(row)?.getCell((short)col)
}

// ----------------------------
// エンティティ情報を取得する
//    今回は必要な情報のセル座標はgivenなものとする(キーで探索などはしない)
def table = new Table(
    physicalName: cell(2,0).stringCellValue,
    logicalName: cell(2,1).stringCellValue
)
for (i in 6..100) {
    // physicalNameが空の場合はそこで打ち切る
    if (!cell(i, 0)?.stringCellValue) break

    table.columns << new Column(
        physicalName: cell(i, 0)?.stringCellValue,
        logicalName:  cell(i, 1)?.stringCellValue,
        type:         cell(i, 2)?.stringCellValue.toUpperCase(),
        length:       cell(i, 3)?.numericCellValue.intValue(), // Double -> int
        notNull:      cell(i, 4)?.booleanCellValue
    )
}
//println table

// ----------------------------
// ソースファイルを出力する

// 先頭を大文字化するクロージャ(メソッドでも良いが、groovyっぽくクロージャで)
def capitalize = { text -> text[0].toUpperCase() + text.substring(1) }

// カラム情報からフィールド・メソッド情報を生成する
def imports = []
def fields = []
def methods = []
table.columns.each { column ->
    def propertyName = column.physicalName
    def capitalizedName = capitalize(propertyName)

    // typeごとに実装(現状は、今使用中の方のみ対応している)
    def javaType
    if (column.type == "VCHAR") {
        javaType = "String"
    }
    else if (column.type == "DATE") {
        imports << "import java.util.Date;"
        javaType = "Date"
    }

    fields  << "private $javaType $propertyName;"
    methods << "public $javaType get$capitalizedName() { return $propertyName; }"
    methods << "public void set$capitalizedName($javaType $propertyName) { this.$propertyName = $propertyName; }"
}

// 出力する
def indent = "    "
def sep = System.getProperty('line.separator')
def className = table.physicalName
new File(className + ".java").withWriter { writer ->
    imports.each { writer << it << sep }
    writer << sep
    writer << "public class $table.physicalName {" << sep
    fields.each  { writer << indent << it << sep }
    writer << sep
    methods.each { writer << indent << it << sep }
    writer << "}" << sep
}
