package org.selenium;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ExcelReadAdWrite {
	public static void main(String[] args) throws IOException {
	File f= new File("C:\\Users\\user\\eclipse-workspace\\Selenium_Feb\\src\\test\\resources\\Excel.xlsx");
    FileOutputStream f1= new FileOutputStream(f);
    Workbook book= new HSSFWorkbook();
     Sheet sheet= book.createSheet("tools");
     for(int i=0;i<10;i++) {
    	Row row= sheet.createRow(i);
    for(int j=0;j<3;j++) {
    	Cell cell= row.createCell(j);
    	cell.setCellValue(1);
    	
    }
     }   
     book.write(f1);
     book.close();
     System.out.println("test");
	}
}
