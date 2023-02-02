package day5;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRead {

	public static void main(String[] args) throws IOException {

		FileInputStream fis1= new FileInputStream("src\\test\\resources\\testdata\\excels\\LoginData.xlsx");
		
		XSSFWorkbook wb= new XSSFWorkbook(fis1);
		
		XSSFSheet  ws=wb.getSheet("Sheet1");
		
		//iterator on Sheet
		
		Iterator<Row>  rows=ws.iterator();
		
		Row row=null;
		
		Cell cell=null;
		
		while(rows.hasNext())
		{
			row=rows.next();
			
			Iterator<Cell> cells=row.iterator();
			
			while(cells.hasNext())
			{
				
				cell=cells.next();
				
				if(cell.getCellType()==CellType.STRING)
				{
					System.out.println(cell.getStringCellValue()+" is a String");
				}
				
				else if(cell.getCellType()==CellType.NUMERIC)
				{
					System.out.println(cell.getNumericCellValue()+" is a Number");
				}
				
				else if(cell.getCellType()==CellType.BOOLEAN)
				{
					System.out.println(cell.getBooleanCellValue()+" is a Boolean");
				}
				
				
			}
			
			
		}
		
		
	
		
		
	}

}
