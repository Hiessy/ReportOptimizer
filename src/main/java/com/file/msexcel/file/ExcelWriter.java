package com.file.msexcel.file;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class ExcelWriter {

	public static void main(String[] args) throws Exception {
		writeIntoExcel("src//resources//reporte_test_cc.xls");
	}

	public static void writeIntoExcel(String file) throws FileNotFoundException, IOException, EncryptedDocumentException, InvalidFormatException {

		InputStream inputStream = new FileInputStream(file);

		@SuppressWarnings("resource")
		HSSFWorkbook myExcelBook = new HSSFWorkbook(inputStream);
		//Workbook wb = WorkbookFactory.create(inputStream);
		HSSFSheet myExcelSheet = myExcelBook.getSheet("Sheet1");
		//Sheet sheet = myExcelBook.getSheet("Sheet1");
		Row row = myExcelSheet.getRow(1);
		if(row == null)
			row = myExcelSheet.createRow(1);
		Cell cell = row.getCell(3);
		if (cell == null)
			cell = row.createCell(3);
		cell.setCellType(Cell.CELL_TYPE_STRING);
		cell.setCellValue("a test");

		// Write the output to a file
		FileOutputStream fileOut = new FileOutputStream(file);
		myExcelBook.write(fileOut);
		fileOut.close();
	}
}
