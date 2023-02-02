package day4;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRead {

	public static void main(String[] args) throws IOException {

		FileInputStream fis1= new FileInputStream("src\\test\\resources\\testdata\\excels\\LoginData.xlsx");
		
		XSSFWorkbook wb= new XSSFWorkbook(fis1);
		
		XSSFSheet  ws=wb.getSheet("Sheet1");
		
		
		int lastRowNum=ws.getLastRowNum();
		
		System.out.println("lastRowNum :"+lastRowNum);
		
		int physicalNumberOfRows=ws.getPhysicalNumberOfRows();
		
		System.out.println("physicalNumberOfRows :"+physicalNumberOfRows);
		
		//int lastCellNum=ws.getRow(0).getLastCellNum();
		
		int physicalNumberOfCell=ws.getRow(0).getPhysicalNumberOfCells();
		
		//System.out.println("lastCellNum :"+lastCellNum);
		
		System.out.println("physicalNumberOfCell :"+physicalNumberOfCell);
		
		Row row=null;
		
		for(int i=0;i<=lastRowNum;i++)
		{
			
			row=ws.getRow(i);
			
			if(!(row==null))
			{
			int lastCellNum=row.getLastCellNum();
			
			for(int j=0;j<lastCellNum;j++)
			{
				if(row.getCell(j).getCellType() == CellType.STRING)
				{
					System.out.println(row.getCell(j).getStringCellValue()+" is a String");
				}
				else if(row.getCell(j).getCellType() == CellType.NUMERIC)
				{
					System.out.println(row.getCell(j).getNumericCellValue()+" is a Number");
				}
				else if(row.getCell(j).getCellType() == CellType.BOOLEAN)
				{
					System.out.println(row.getCell(j).getBooleanCellValue()+" is a Boolean");
				}
			}
			}
		}
		

	}

}
