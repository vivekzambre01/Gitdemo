package Utility;

import java.io.IOException;

public class ExcelUtilsDemo {

	public static void main(String[] args) throws IOException {
		String projectpath = System.getProperty("user.dir");
		ExcelUtills excel = new ExcelUtills(projectpath+"/Excel/data.xlsx", "sheet1");
		
		excel.getcelldataNumber(1, 1);
		excel.getcelldatasTring(0, 0);
	}

}
