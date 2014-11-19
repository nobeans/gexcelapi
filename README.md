GExcelAPI
=========

GExcelAPI is a thin Groovy-ish wrapper library of not JExcelAPI but Apache POI.


Getting Started
---------------

It's difficult to read and write when using Apache POI directly.
Especially, an identification of a cell to use an index is too complicated and ugly.

```groovy
File inputFile = ...
def book = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(inputFile)))
def sheet = book.getSheetAt(0) // 1st sheet
println "A1: " + sheet.getRow(0)?.getCell((short) 0)
println "A2: " + sheet.getRow(1)?.getCell((short) 0)
println "B1: " + sheet.getRow(0)?.getCell((short) 1)
```

By using GExcelAPI, you can write the above sample like this:

```groovy
File inputFile = ...
def book = GExcel.open(inputFile)
def sheet = book[0] // 1st sheet
println "A1: " + sheet.A1.value
println "A2: " + sheet.A2.value
println "B1: " + sheet.B1.value
```


How to get
-----------

### Grape

```groovy
@GrabResolver(name="bintray", root="http://dl.bintray.com/nobeans/maven")
@GrabConfig(systemClassLoader=true) // necessary if you invoke it by GroovyServ
@Grab("org.jggug.kobo:gexcelapi:0.4")
import org.jggug.kobo.gexcelapi.GExcel

// example...
def book = GExcel.open(args[0])
def sheet = book[0]
println sheet.A1.value
```


Test as Documentation
---------------------

* <https://github.com/nobeans/gexcelapi/blob/master/src/test/groovy/org/jggug/kobo/gexcelapi/GExcelTest.groovy>


Code Status
-----------

[![Build Status](https://travis-ci.org/nobeans/gexcelapi.svg?branch=master)](https://travis-ci.org/nobeans/gexcelapi)


License
-------

GExcelAPI is released under the [Apache 2.0 License](http://www.apache.org/licenses/LICENSE-2.0)
