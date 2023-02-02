package day6;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWrite {

	public static void main(String[] args) throws IOException {
		
		FileInputStream fis1= new FileInputStream("src\\test\\resources\\testdata\\excels\\LoginData.xlsx");
		
		XSSFWorkbook wb= new XSSFWorkbook(fis1);
		
		XSSFSheet  ws=wb.getSheet("Sheet1");
		
		//case 1 : Modifying existing cell
		
		//ws.getRow(0).getCell(1).setCellValue("xyz");
		
		//case 2: Create a cell in existing row
		
		//ws.getRow(2).createCell(3).setCellValue("Passed");
		
		//case 3: Creating a cell in new Row
		
		ws.createRow(5).createCell(0).setCellValue(456);
		
		FileOutputStream  fos=new FileOutputStream("src\\test\\resources\\testdata\\excels\\LoginData.xlsx");
		
		wb.write(fos);
		
		wb.close();
		
		System.out.println("End of Program");

	}

}
