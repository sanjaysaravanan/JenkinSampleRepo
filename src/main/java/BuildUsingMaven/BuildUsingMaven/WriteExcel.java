package BuildUsingMaven.BuildUsingMaven;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class WriteExcel 
{
	@Test
    public void readExcel() throws IOException
    {
    	System.out.println("WriteExcel File");
    	XSSFWorkbook xl = new XSSFWorkbook("./Data/DataForBuild.xlsx");
    	XSSFSheet sheet = xl.getSheetAt(0);
    	int count = sheet.getLastRowNum();
    	System.out.println(count);
    	int colCount = sheet.getRow(0).getLastCellNum();
    	System.out.println(colCount);
    	
    	for(int i = 1; i <= count ; i++){
    		XSSFRow row = sheet.getRow(i);
    		for (int j = 0; j < colCount; j++) {
				XSSFCell cell = row.getCell(j);
				System.out.println(cell.getStringCellValue());
			}
    	}
    }
}
