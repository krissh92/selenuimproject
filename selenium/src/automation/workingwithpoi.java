package automation;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

public class workingwithpoi {

	
ChromeDriver driver = new org.openqa.selenium.chrome.ChromeDriver();
	
	@BeforeTest
	public void open() throws InterruptedException {
		driver.get("https://www.google.com/");
		//driver.manage().window().maximize();
	
		Thread.sleep(3000);
	}
	
	@Test
	public void data() throws InterruptedException, IOException {
		
		//FileInputStream fis= new FileInputStream("‪‪C:\\Users\\kn025314\\Desktop\\Book1.xlsx");
		 FileInputStream fis = new FileInputStream("C:\\Users\\kn025314\\Desktop\\Book1.xlsx");
		XSSFWorkbook wbo=new XSSFWorkbook(fis);
		XSSFSheet wso= wbo.getSheet("Sheet1");
		Row r;
		String value =driver.findElement(By.linkText("Gmail")).getText();
		r=wso.createRow(0);
		r.createCell(0).setCellValue(value);
		//FileOutputStream fos= new FileOutputStream("‪‪C:\\Users\\kn025314\\Desktop\\Book1.xlsx");
		FileOutputStream fos= new FileOutputStream("C:\\\\Users\\\\kn025314\\\\Desktop\\\\Book1.xlsx");
		wbo.write(fos);
	}

	

}
