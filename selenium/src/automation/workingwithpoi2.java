package automation;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class workingwithpoi2 {

	public static void main(String[] args) throws IOException  {
		// TODO Auto-generated method stub
		FileInputStream fis = new FileInputStream("C:\\Users\\kn025314\\Desktop\\Book1.xlsx");
		//FileInputStream fis= new FileInputStream("‪‪C:\\Users\\kn025314\\Book1.xlsx");
		XSSFWorkbook wbo =new XSSFWorkbook(fis);
		XSSFSheet wso= wbo.getSheet("Sheet1");
		Row r;
		//String value =driver.findElement(By.linkText("Gmail")).getText();
		
		r=wso.createRow(0);
		r.createCell(0).setCellValue("testing");
		FileOutputStream fos= new FileOutputStream("C:\\\\Users\\\\kn025314\\\\Desktop\\\\Book1.xlsx");
		wbo.write(fos);
		
 
	}

}
