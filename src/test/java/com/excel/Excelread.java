package com.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excelread {
	
	public static void main(String[] args) throws IOException {
		
	
	File file = new File("C:\\Users\\Arun\\eclipse-workspace3\\DataDrivenFrame\\src\\test\\resources\\DataDriven Excel.xlsx");

	FileInputStream fis = new FileInputStream(file);
	
	Workbook w = new XSSFWorkbook(fis); 
	
	Sheet sheet = w.getSheet("Sheet1");
	
	Row row = sheet.getRow(1);
	Cell cell = row.getCell(0);
	System.out.println(cell);
	
	CellType cellType = cell.getCellType();
	System.out.println(cellType);
	
//	for(int i=0;i<sheet.getPhysicalNumberOfRows();i++) {
//		Row row = sheet.getRow(i);
//		for(int j=0;j<row.getPhysicalNumberOfCells();j++) {
//			Cell cell = row.getCell(j);
//			System.out.println(cell);
//		
//			
//		}
	}
		
	
	
	
}
	
	
