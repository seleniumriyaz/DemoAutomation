package day3;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class ExcelRead {
	
	
	@Test
	public void excelReadTest() throws IOException
	{
		
		FileInputStream fis1= new FileInputStream("src\\test\\resources\\testdata\\excels\\LoginData.xlsx");
		
		XSSFWorkbook wb= new XSSFWorkbook(fis1);
		
		XSSFSheet  ws=wb.getSheet("Sheet1");
		
		
		int lastRowNum=ws.getLastRowNum();
		
		System.out.println("lastRowNum :"+lastRowNum);
		
		
	}

}
