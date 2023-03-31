package excelFileRead;

import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MyExcelReadCode 
{
	static FileInputStream f;
	static XSSFWorkbook wb;
	static XSSFSheet sh;
	
//for read Name in the excel sheet
	public static String stringDataRead(int row, int col) throws IOException
	{
		f=new FileInputStream("C:\\Users\\A-10680\\Desktop\\Test Java file\\Test Data.xlsx");
		wb=new XSSFWorkbook(f);
		sh=wb.getSheet("Sheet1");
		Row r= sh.getRow(row);
		Cell c=r.getCell(col);
		return c.getStringCellValue();	
	}
	
//for read Age value in the excel sheet
	public static String integerDataRead(int row, int col) throws IOException
	{
		f=new FileInputStream("C:\\Users\\A-10680\\Desktop\\Test Java file\\Test Data.xlsx");
		wb=new XSSFWorkbook(f);
		sh=wb.getSheet("Sheet1");
		Row r= sh.getRow(row);
		Cell c=r.getCell(col);
		int a=(int) c.getNumericCellValue();		//type conversion
		return String.valueOf(a);	
	}
		
	
	
	//We can write in this way also-need to change string to int in main method also.
	
/*	public static int integerDataRead(int row, int col) throws IOException
	{
		f=new FileInputStream("C:\\Users\\A-10680\\Desktop\\Test Java file\\Test Data.xlsx");
		wb=new XSSFWorkbook(f);
		sh=wb.getSheet("Sheet1");
		Row r= sh.getRow(row);
		Cell c=r.getCell(col);
		//int a=(int) c.getNumericCellValue();		//type conversion
		return (int) c.getNumericCellValue();
	}*/
	
}
