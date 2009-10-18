@Grab(group='org.apache.poi', module='poi', version='3.5-beta3')
import org.apache.poi.hssf.usermodel.*
import org.apache.poi.poifs.filesystem.*

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

def inputFile = new File(args[0])
def book = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(inputFile)))
def sheet = book.getSheetAt(0) 
def cell = { row, col ->
    sheet.getRow(row)?.getCell((short)col)
}

def table = new Table(
    physicalName: cell(2,0).stringCellValue,
    logicalName: cell(2,1).stringCellValue
)
for (i in 6..100) {
    
    if (!cell(i, 0)?.stringCellValue) break

    table.columns << new Column(
        physicalName: cell(i, 0)?.stringCellValue,
        logicalName:  cell(i, 1)?.stringCellValue,
        type:         cell(i, 2)?.stringCellValue.toUpperCase(),
        length:       cell(i, 3)?.numericCellValue.intValue(), 
        notNull:      cell(i, 4)?.booleanCellValue
    )
}

def capitalize = { text -> text[0].toUpperCase() + text.substring(1) }

def imports = []
def fields = []
def methods = []
table.columns.each { column ->
    def propertyName = column.physicalName
    def capitalizedName = capitalize(propertyName)

    
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
