package org.selenium;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import io.github.bonigarcia.wdm.WebDriverManager;

public class FlipkartRW {
	public static void main(String[] args) throws IOException {
		WebDriverManager.chromedriver().setup();
		WebDriver driver= new ChromeDriver();
		driver.manage().window().maximize();
		driver.get("https://www.amazon.com");
		driver.manage().timeouts().implicitlyWait(10,TimeUnit.SECONDS);
		File f= new File("C:\\Users\\user\\eclipse-workspace\\Selenium_Feb\\src\\test\\resources\\Excel1.xlsx");
		 FileOutputStream f1= new FileOutputStream(f);
		 Workbook book= new HSSFWorkbook();
		 Sheet sheet= book.createSheet("Nokia");
		 WebElement search= driver.findElement(By.id("twotabsearchtextbox"));
			search.sendKeys("nokia");
			WebElement searchIcon= driver.findElement(By.id("nav-search-submit-button"));
			searchIcon.click();
		List<WebElement> nokia= driver.findElements(By.xpath("//div[@class='a-section a-spacing-small a-spacing-top-small']"));
		for(int i=0; i<nokia.size();i++) {
			WebElement ele=nokia.get(i);
			Row row= sheet.createRow(i);
			Cell cell= row.createCell(0);
			cell.setCellValue(ele.getText());
		}
		 book.write(f1);
	     book.close();
	     System.out.println("*******************");
	     
	     FileInputStream f2= new FileInputStream(f);
	     Workbook book1= new HSSFWorkbook(f2);
	     Sheet sheet1= book1.getSheet("nokia");
	     for(int i=0;i<nokia.size();i++) {
	     Row row1= sheet1.getRow(i);
	     Cell cell1 =row1.getCell(0);
	     
	     try {
	    	 String cellValue= cell1.getStringCellValue();
	    	 System.out.println(cellValue);
	     }
	     catch(Exception e) {
	    	 double cellValue= cell1.getNumericCellValue();
	    	 System.out.println(cellValue);
	     }
	    
	     }
	     System.out.println("*******************");
	}
	

}
