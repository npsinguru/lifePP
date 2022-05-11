package DataDriven;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class GetDataFromExcel 
{
	public void GetData(String testCase) throws IOException
	{
		String sheetName = "Requisitions";
		String testCaseName = "";
		String fetchfield = "";
		String filePath="C:\\Users\\nsinguru\\OneDrive - DXC Production\\Selenium Project\\lifePP\\TestData\\ZCHE Test Data.xlsx";
		FileInputStream fis = new FileInputStream(filePath);
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		int sheets = workbook.getNumberOfSheets();
		for(int i=0; i<sheets;i++)
		{
			if(workbook.getSheetName(i).equalsIgnoreCase(sheetName))
			{
				XSSFSheet reqSheet = workbook.getSheetAt(i);
				Iterator<Row> rows = reqSheet.iterator();
				while(rows.hasNext())
				{
					Row allrows = rows.next();
				}
			}
		}
		
	}
	
		
	
	
	
}