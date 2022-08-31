package org.Excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DDRead {
	
	public static void main(String[] args) throws IOException  {
		
		File f = new File("C:\\Users\\Arun\\eclipse-workspace3\\DataDrivenFrame\\src\\test\\resources\\ReadExcel2.xlsx");
		
		FileInputStream fis = new FileInputStream(f);
		
		Workbook w = new XSSFWorkbook(fis);
		
		Sheet s = w.getSheet("Sheet1");
		
		Row row = s.getRow(3);
		Cell cell = row.getCell(3);
		System.out.println(cell);
		
//		String stringCellValue = cell.getStringCellValue();
//		System.out.println(stringCellValue);
		
		
		
	}

}
