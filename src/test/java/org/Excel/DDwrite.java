package org.Excel;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DDwrite {

	public static void main(String[] args) throws IOException {


		File f = new File("C:\\Users\\Arun\\eclipse-workspace3\\DataDrivenFrame\\src\\test\\resources\\DD.xlsx");

		Workbook w = new XSSFWorkbook();

		Sheet sheet = w.createSheet("bala");

		Row row = sheet.createRow(3);

		Cell cell = row.createCell(2);
		cell.setCellValue("xxyy");
		
		FileOutputStream fos = new FileOutputStream(f);
		
		w.write(fos);
		
		System.out.println("done");


	}

}
