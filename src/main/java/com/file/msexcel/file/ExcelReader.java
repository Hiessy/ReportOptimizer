package com.file.msexcel.file;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.file.msexcel.util.ObservacionesSplitter;

public class ExcelReader {

	public static void main(String[] args) throws Exception {
		List<String> cellObservacionesFiltered = readFromExcel("src//resources//reportVASMCSS_1415903240717.xls");
		// writeIntoExcel("src//resources//reportVASMCSS_test.xls");

	}

	/**
	 * Java method to read dates from Excel file in Java. This method read value
	 * from .XLS file, which is an OLE format.
	 * 
	 * @param file
	 * @throws Exception
	 */

	@SuppressWarnings("resource")
	public static List<String> readFromExcel(String file) throws Exception {

		FileInputStream inputStream = new FileInputStream(file);
		HSSFWorkbook myExcelBook = new HSSFWorkbook(inputStream);
		HSSFSheet myExcelSheet = myExcelBook.getSheet("report_VASMCSS ");

		List<String> headers = new ArrayList<String>();
		int ObservacionColIndex = 0;

		for (int i = 0; i < myExcelSheet.getRow(0).getLastCellNum(); i++) {

			if (!myExcelSheet.getRow(0).getCell(i).getStringCellValue().equals("Observaciones")) {
				headers.add(myExcelSheet.getRow(0).getCell(i).getStringCellValue());
			} else {
				ObservacionColIndex = i;
			}

		}

		System.out.println("Showing relevant headers for excel file " + file + ":" + headers);

		if (ObservacionColIndex == 0)
			throw new Exception("La hoja de Excel no tiene la columna de observaciones");

		List<String> cellObservaciones = new ArrayList<String>();
		List<String> result = new ArrayList<String>();
		String cellValueString;

		// filtra resultados para la celda i debajo del campo "Observaciones"
		for (int i = 1; i <= myExcelSheet.getLastRowNum(); i++) {

			cellValueString = myExcelSheet.getRow(i).getCell(ObservacionColIndex).getStringCellValue();
			cellObservaciones = ObservacionesSplitter.split(cellValueString, "\\|");

			for (int n = 0; n < cellObservaciones.size(); n++) {
				for (int m = 0; m < headers.size(); m++) {
					if (!headers.get(m).equals("")) {
						if (cellObservaciones.get(n).contains(headers.get(m))) {
							result.add(cellObservaciones.get(n));
							break;
						}
					}
				}

			}
			// TODO -- need to clear list before adding more and have to find a
			// way to return posiotn 0 cell 1 and so on.
			// System.out.println("Showing none filter results for row " + i +
			// ": " + resultObservaciones);
			// System.out.println("Showing filter results for row " + i +": " +
			// resultObservacionesFiltered);
			// myExcelSheet.
			if (!cellObservaciones.equals(result)) {
				System.out.println("Observaciones for row " + (i + 1) + " is diferent -- " + result);
			}

			result.clear();
		}

		/*
		 * HSSFRow row = myExcelSheet.getRow(0);
		 * 
		 * if(row.getCell(0).getCellType() == HSSFCell.CELL_TYPE_STRING){ String
		 * name = row.getCell(0).getStringCellValue();
		 * System.out.println("name : " + name); }
		 */

		myExcelBook.close();
		return result;
	}

}