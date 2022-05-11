package DataDriven;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.util.SystemOutLogger;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;





public class ExcelDriven {
	
	public static void main (String arg[]) throws IOException {
		
	
	
	//public ArrayList<String> ExcelGetData() throws IOException
	{
		String sheetName = "Requisitions";
		String testCaseField = "ID";
		String testCaseId = "TC 0006 EWP - Create_Withdraw Requisition";
		String requiredField = "DefaultAccountAssignmentRequisition";
		
		
		ArrayList<String> excelExpValues = new ArrayList<>();
		ArrayList arr = new ArrayList();
		
		FileInputStream fis = new FileInputStream("C:\\Users\\nsinguru\\OneDrive - DXC Production\\Selenium Project\\lifePP\\TestData\\ZCHE Test Data.xlsx");
		XSSFWorkbook excelWorkBook = new XSSFWorkbook(fis);
		int noOfSheets = excelWorkBook.getNumberOfSheets();
		
		for(int i=0; i<noOfSheets;i++)
		{
			if(excelWorkBook.getSheetName(i).equalsIgnoreCase(sheetName))
			{
				XSSFSheet reqSheet = excelWorkBook.getSheetAt(i);	
				Iterator<Row> allRows=reqSheet.iterator();
				Row firstRow = allRows.next();
				Iterator<Cell> reqCell = firstRow.cellIterator();
				int column1 = 0;
				int column2 = 0;
				int row=0;
				
				while(reqCell.hasNext())
				{
					Cell reqCellValue=reqCell.next();
					if(reqCellValue.getStringCellValue().equalsIgnoreCase(testCaseField))
					{
						int colOne = column1;
						System.out.println("Required Column one is = "+colOne);
						break;
					
					}
					column1++;
					
				}
				while(allRows.hasNext())
				
				{	
					Row rows=allRows.next();
											
					if(rows.getCell(column1).getStringCellValue().equalsIgnoreCase(testCaseId))
					{
						Iterator<Cell> reqCellValue = rows.cellIterator();
						while(reqCellValue.hasNext())
						{	
							Cell cellValues = reqCellValue.next();
							if(cellValues.getCellType()==CellType.STRING)
							{
								
								excelExpValues.add(cellValues.getStringCellValue());
								System.out.println(cellValues.getStringCellValue());
								
							}
							else
							{
								excelExpValues.add(NumberToTextConverter.toText(cellValues.getNumericCellValue()));
								System.out.println(NumberToTextConverter.toText(cellValues.getNumericCellValue()));
								
							}
							
						
						}
					}
							
					
				}
				
				
			}
			
		}
		//return excelExpValues;
		
		
	}
	}
	
}
	
