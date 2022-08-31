package com.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {

	public static void main(String[] args) throws IOException {
		String res;
		File f = new File(
				"C:\\Users\\Arun\\eclipse-workspace3\\DataDrivenFrame\\src\\test\\resources\\ReadExcel2.xlsx");

		FileInputStream fis = new FileInputStream(f);

		Workbook w = new XSSFWorkbook(fis);

		Sheet sheet = w.getSheet("sheet1");

		Row row = sheet.getRow(0);
		Cell cell = row.getCell(1);
		System.out.println(cell);
		Row row1 = sheet.getRow(7);

		Cell cell1 = row1.getCell(1);
		String value = cell.getStringCellValue();
		
		
		if (value.contains("gwfiu")){
			cell.setCellValue("bala");
		}
			
		FileOutputStream foss = new FileOutputStream(f);
		w.write(foss);
		System.out.println("completed");
		
		
		
		
		
		
		
		

//		for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
//			Row row2 = sheet.getRow(i);
//			for (int j = 0; j < row2.getPhysicalNumberOfCells(); j++) {
//				Cell cell2 = row2.getCell(j);
//				System.out.println(cell2);
//				CellType cellType = cell.getCellType();
//				System.out.println(cellType);
//
//				switch (cellType) {
//				case STRING:
//					res = cell2.getStringCellValue();
//					System.out.println(res);
//					break;
//
//				case NUMERIC:
//					double d = cell2.getNumericCellValue();
//					long l = (long) d;
//					String valueOf = String.valueOf(l);
//					System.out.println(valueOf);
//					break;
//
//				}
//
//			}
//
//		}
//
//		CellType cellType = cell.getCellType();
//		System.out.println(cellType);
//		
//		CellType cellType1 = cell1.getCellType();
//		System.out.println(cellType1);

	}

}
