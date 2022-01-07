package operations;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.*;

public class ExcelOperations {

	String excelFilePath = "I:\\Selenium\\Data Operation\\src\\main\\resources\\ExcelFiles\\countries.xlsx";
	FileInputStream inputStream;
	FileOutputStream outputStream;
	Iterator iterator;

	XSSFWorkbook workbook;
//	XSSFSheet sheet;
	XSSFSheet sheet;
	XSSFRow row;
	XSSFCell cell;

	// Reading all data from the excel sheet using for loop
	public void readingExcelForLoop() throws IOException {

		inputStream = new FileInputStream(excelFilePath);
		workbook = new XSSFWorkbook(inputStream);
//		sheet = workbook.getSheet("Sheet1");
		sheet = workbook.getSheetAt(0);

		int rows = sheet.getLastRowNum();
		int cols = sheet.getRow(1).getLastCellNum();

		for (int r = 0; r <= rows; r++) {
			row = sheet.getRow(r);

			for (int c = 0; c < cols; c++) {
				cell = row.getCell(c);

				switch (cell.getCellType()) {

				case STRING:
					System.out.print(cell.getStringCellValue());
					break;
				case NUMERIC:
					System.out.print(cell.getNumericCellValue());
					break;
				case BOOLEAN:
					System.out.print(cell.getBooleanCellValue());
					break;
				}
				System.out.print("  |  ");
			}
			System.out.println();
		}
	}

	// Reading all data from the excel sheet using iterator
	public void readingExcelIterator() throws IOException {

		inputStream = new FileInputStream(excelFilePath);
		workbook = new XSSFWorkbook(inputStream);
//		sheet = workbook.getSheet("Sheet1");
		sheet = workbook.getSheetAt(0);
		iterator = sheet.iterator();

		while (iterator.hasNext()) {
			row = (XSSFRow) iterator.next();
			Iterator cellIterator = row.cellIterator();

			while (cellIterator.hasNext()) {
				cell = (XSSFCell) cellIterator.next();

				switch (cell.getCellType()) {

				case STRING:
					System.out.print(cell.getStringCellValue());
					break;
				case NUMERIC:
					System.out.print(cell.getNumericCellValue());
					break;
				case BOOLEAN:
					System.out.print(cell.getBooleanCellValue());
					break;
				}
				System.out.print("  |  ");
			}
			System.out.println();
		}
	}

	// Writing data to excel sheet
	public void writingExcel() throws IOException {

		workbook = new XSSFWorkbook();
		sheet = workbook.createSheet("EmpInfo");

		Object empdata[][] = { { "EmpID", "Name", "Job" }, { 101, "David", "Engineer" }, { 102, "Smith", "Manager" },
				{ 103, "Scott", "Engineer" } };

		/*
		 * // Using for loop int rows = empdata.length; int cols = empdata[0].length;
		 * 
		 * System.out.println(rows); System.out.println(cols);
		 * 
		 * for (int r = 0; r < rows; r++) {
		 * 
		 * row = sheet.createRow(r); for (int c = 0; c < cols; c++) {
		 * 
		 * cell = row.createCell(c); Object value = empdata[r][c];
		 * 
		 * if (value instanceof String) cell.setCellValue((String) value); if (value
		 * instanceof Integer) cell.setCellValue((Integer) value); if (value instanceof
		 * Boolean) cell.setCellValue((Boolean) value); } }
		 */

		// Using foreach loop
		int rowCount = 0;
		for (Object emp[] : empdata) {
			row = sheet.createRow(rowCount++);
			int cellCount = 0;
			for (Object value : emp) {
				cell = row.createCell(cellCount++);

				if (value instanceof String)
					cell.setCellValue((String) value);
				if (value instanceof Integer)
					cell.setCellValue((Integer) value);
				if (value instanceof Boolean)
					cell.setCellValue((Boolean) value);
			}
		}

		String outputFilePath = "I:\\Selenium\\Data Operation\\src\\main\\resources\\ExcelFiles\\employee.xlsx";
		outputStream = new FileOutputStream(outputFilePath);
		workbook.write(outputStream);
		outputStream.close();

		System.out.println("employee.xlsx file is written successfully");
	}

}
