package org.jggug.kobo.gexcelapi

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook

import groovy.util.GroovyTestCase

class GExcelEditTest extends GroovyTestCase {

	def sampleFile = "build/resources/test/tasklist.xls"
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
		//outputFile.delete()
	}

	void testWrite() {
		outputFile.withOutputStream { book.write(it) }
		Workbook targetBook = GExcel.open(outputFile)
		Sheet targetSheet = book[0]

		assert outputFile.exists()
		assert sheet.A1 == targetSheet.A1
		assert sheet.A2 == targetSheet.A2
		assert sheet.B1 == targetSheet.B1
		assert sheet.B2 == targetSheet.B2
	}

	void testFindEmptyRow1() {
		Row emptyRow = sheet.findEmptyRow('A2');
		assert emptyRow != null
		assert emptyRow.label == "11"  // labelプロパティは、Excel行番号と一致
		assert emptyRow.getCell(0)?.value == null
	}

	void testFindEmptyRow2() {
		Row emptyRow = sheet.findEmptyRow('B3');
		assert emptyRow != null
		assert emptyRow.label == "9"  // labelプロパティは、Excel行番号と一致
		assert emptyRow.getCell(1)?.value == null
	}

	void testAddRow() {
		Row emptyRow = sheet.findEmptyRow('A2');
		emptyRow.createCell(0).value = "100"
		emptyRow.createCell(1).value = "test"

		outputFile.withOutputStream { book.write(it) }
		Workbook targetBook = GExcel.open(outputFile)
		Sheet targetSheet = targetBook[0]
		assert targetSheet.A11.value == "100"
		assert targetSheet.B11.value == "test"
	}

	void testFindByCellValue1() {
		Row resultRow = sheet.findByCellValue('F4', '田中');
		assert resultRow != null
		assert resultRow.label == "7"  // labelプロパティは、Excel行番号と一致
		assert resultRow.F_.value == "田中"
		assert resultRow.getCell(5) == sheet.F7
	}
	
	void testFindByCellValue2() {
		//前方一致
		Row resultRow = sheet.findByCellValue('C4', '機能');
		assert resultRow != null
		assert resultRow.label == "4"  // labelプロパティは、Excel行番号と一致
		assert resultRow.C_.value == "機能追加"
		assert resultRow.getCell(2) == sheet.C4
	}

	void testFindByCellValue3() {
		//後方一致
		Row resultRow = sheet.findByCellValue('H6', '済');
		assert resultRow != null
		assert resultRow.label == "6"  // labelプロパティは、Excel行番号と一致
		assert resultRow.H_.value == "対応済"
		assert resultRow.getCell(7) == sheet.H6
	}
	
	void testFindByCellValue4() {
		//数値検索
		Row resultRow = sheet.findByCellValue('I4', 20);
		assert resultRow != null
		assert resultRow.label == "5"  // labelプロパティは、Excel行番号と一致
		assert resultRow.I_.value == 20
		assert resultRow.getCell(8) == sheet.I5
	}
	
	void testFindAllByCellValue1() {
		def resultRows = sheet.findAllByCellValue('D4', 'B');
		assert resultRows[0].label == "5"
		assert resultRows[0].D_.value == "B"
		assert resultRows[1].label == "6"
		assert resultRows[1].D_.value == "B"
	}
	
	void testFindAllByCellValue2() {
		def resultRows = sheet.findAllByCellValue('C4', '変更');
		assert resultRows[0].label == "5"
		assert resultRows[0].C_.value == "仕様変更"
		assert resultRows[1].label == "7"
		assert resultRows[1].C_.value == "仕様変更"
	}
	
	void testFindAllByCellValue3() {
		//数値検索
		def resultRows = sheet.findAllByCellValue('I4', 10);
		assert resultRows[0].label == "7"
		assert resultRows[0].I_.value == 10
		assert resultRows[1].label == "8"
		assert resultRows[1].I_.value == 10
	}
}
