package org.cts;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Rocks {
	public static void main(String[] args) throws Throwable {
		File loc = new File("A:\\Isaacraja\\Space\\TestData\\raaa.xlsx");
		Workbook w = new XSSFWorkbook();
		Sheet s = w.createSheet("Isaac123");
		Row r = s.createRow(1);
		Cell c = r.createCell(3);
		c.setCellValue("raja");
		FileOutputStream o = new FileOutputStream(loc);
		w.write(o);
		System.out.println("Written Successfully");
		System.out.println("Written Successfully");
		System.out.println("Written Successfully");
		
	}
	

}
