package operations;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.*;

public class ExcelOperations {

	String excelFilePath = "I:\\Selenium\\Data Operation\\src\\main\\resources\\ExcelFiles\\countries.xlsx";
	FileInputStream inputStream;
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

}
