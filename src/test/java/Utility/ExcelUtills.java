package Utility;

import java.io.IOException;

import org.apache.poi.xssf.dev.XSSFSave;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtills {
	
	static  String projectpath;
	static  XSSFWorkbook wb;
	static  XSSFSheet sheet;
	
	
	public ExcelUtills(String excelpath, String sheetpath) throws IOException
	{
		 projectpath = System.getProperty("user.dir");
	    	wb = new XSSFWorkbook(excelpath);
			 sheet= wb.getSheet(sheetpath);
	}
public static void main(String[] args) {
	
	//getrowcount();
	getcelldatasTring(0,0);
	getcelldataNumber(1,1);
}
	
public static void getrowcount() 
{

	try {
       
		int rowcount = sheet.getPhysicalNumberOfRows();
		System.out.println("no of rows coloum"+rowcount);
	} catch (Exception e) {
		System.out.println(e.getMessage());
		System.out.println(e.getCause());
		e.printStackTrace();
	}
	
}
	public static void getcelldatasTring(int rowNum, int colNum)
	{
		 try {
			
			String celldata = sheet.getRow(rowNum).getCell(colNum).getStringCellValue();
			System.out.println(celldata);
		} catch (Exception e) {
		
			e.printStackTrace();
			e.getMessage();
			
		}
	}
	public static void getcelldataNumber(int rowNum, int colNum)
	{
		 try {
			
			double celldata = sheet.getRow(rowNum).getCell(colNum).getNumericCellValue();
			System.out.println(celldata);
		} catch (Exception e) {
		
			e.printStackTrace();
			e.getMessage();
			
		}
	}
	
	
	
	
	
	
	
	
}
