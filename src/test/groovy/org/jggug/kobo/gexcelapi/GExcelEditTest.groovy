package org.jggug.kobo.gexcelapi

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook

import groovy.util.GroovyTestCase

class GExcelEditTest extends GroovyTestCase {

    def sampleFile = "build/resources/test/sample.xls"
	def outputPath = "output.xls"
	File outputFile
    Workbook book
    Sheet sheet
	
	void setUp() throws Exception {
        book = GExcel.open(sampleFile)
        sheet = book[0]
        outputFile = new File(outputPath)
	}
	
    void tearDown() {
        book = null
        sheet = null
//		outputFile.delete()
    }

	void testWrite() {
		outputFile.withOutputStream {
			book.write(it)
		}
		Workbook targetBook = GExcel.open(outputFile)
		Sheet targetSheet = book[0]
		
		assert outputFile.exists()
		assert sheet.A1 == targetSheet.A1
		assert sheet.A2 == targetSheet.A2
		assert sheet.B1 == targetSheet.B1
		assert sheet.B2 == targetSheet.B2
	}

	void testFindEmptyRow() {
		Row emptyRow = sheet.findEmptyRow(1);
		assert emptyRow != null
		assert emptyRow.getCell(0) == null
	}
	
	void testAddRow() {
		Row emptyRow = sheet.findEmptyRow(1);
		emptyRow.createCell(0).value = "hello"
		
		outputFile.withOutputStream {
			book.write(it)
		}
		Workbook targetBook = GExcel.open(outputFile)
		Sheet targetSheet = book[0]
		assert targetSheet.A8.value == "hello"
	}
}
