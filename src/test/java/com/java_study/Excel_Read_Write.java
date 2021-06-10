package com.java_study;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel_Read_Write {
	public static void main(String[] args) throws IOException {
		
		
		/*
		 *   
		 *   1. for achieve using java to read write to Excel we using the Apache-POI,
		 * download dependency of poi and poi-ooxml, version is 5.0.0. 
		 *   2.using File class to specifying where the excel is located.
		 *   3.we are opening the input stream==>A file input stream is an input stream for reading data from a File or from a FileDescriptor .
		 *   4.Connecting to the excel workbook.
		 *   5.Connecting to a particular sheet within the workbook
		 *   6.getting the number of utilized rows
		 */
		
		//The File class from the java.io package, allows us to work with files.
		//To use the File class, create an object of the class, 
		//  and specify the filename or directory name:
		
		File file =new File("./src/test/java/com/java_study/Book3.xlsx"); 
		FileInputStream input=new FileInputStream(file);
		XSSFWorkbook book=new XSSFWorkbook(input);
		XSSFSheet sheet =book.getSheet("Sheet1");
		
		int rowCount = sheet.getPhysicalNumberOfRows();
		System.out.println("Number of row is: " + rowCount);
		
		for(int row=1; row< rowCount; row++) {
			String action =sheet.getRow(row).getCell(0).toString();
			if(action.equalsIgnoreCase("Y"));
			System.out.println(sheet.getRow(row).getCell(1));
			
		}
		input.close();
		book.close();
		
		
		
	}

}
