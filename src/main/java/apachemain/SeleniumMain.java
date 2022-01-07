package apachemain;

import java.io.IOException;

import operations.ExcelOperations;

public class SeleniumMain {

	
	public static void main(String[] args) throws IOException {
		
		ExcelOperations excelOperations = new ExcelOperations();
		
//		excelOperations.readingExcelForLoop();
		excelOperations.readingExcelIterator();
	}

}
