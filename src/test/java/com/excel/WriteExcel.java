package com.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcel {
	
	public static void main(String[] args) throws IOException  {
		
		File f = new File("C:\\Users\\Arun\\eclipse-workspace3\\DataDrivenFrame\\src\\test\\resources\\bala.xlsx");
//		FileInputStream fio = new FileInputStream(f);
		Workbook w = new XSSFWorkbook();
		
		
				
		Sheet sheet = w.createSheet("nw");
		
		Row createRow = sheet.createRow(7 );
		
		Cell createCell = createRow.createCell(4);
		
		createCell.setCellValue("bala");
		
	
		
		FileOutputStream fos = new FileOutputStream(f);
		
		w.write(fos);
		System.out.println("Done");

		
		
		
		
		
	}

}
