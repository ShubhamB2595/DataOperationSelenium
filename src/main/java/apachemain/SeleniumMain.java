package apachemain;

import java.io.IOException;

import operations.ExcelOperations;

public class SeleniumMain {

	static ExcelOperations excelOperations = new ExcelOperations();
	
	
	public static void main(String[] args) throws IOException {
		
//		excelOperations.readingExcelForLoop();
		excelOperations.writingExcel();
	}

}
