package com.ReadExcel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReadData 
{
	public void readExcel(String filename,String sheetname) throws IOException
	{
		int arrayexcel[][]=null;
		FileInputStream fis = new FileInputStream(filename);
		
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet = wb.getSheet(sheetname);
		XSSFRow row = sheet.getRow(2);
		XSSFCell cell = row.getCell(2);
		String val = cell.getStringCellValue();
	    System.out.println("The value at index 2,2 is:===>"+val);
	    
	    // Get rowcount
	    int rows = sheet.getLastRowNum();
	    System.out.println("rows:===>"+rows);
	    
	    int rowcount = rows+1;    // increase index of xlsx 
	    System.out.println("The number of rows are:===>"+rowcount);
	    
	    // Get columcount
	    int columns = sheet.getRow(rows).getLastCellNum();
	    System.out.println("The number of columnss are:===>"+columns);
	    
	    arrayexcel = new int[rowcount][columns];
	    
	    for(int i=0; i<rowcount; i++)
	    {
	    	for(int j=0; j<columns; j++)
	    	{
	    		System.out.println(sheet.getRow(i).getCell(j));
	    		
	    		/* DataFormatter dataformat = new DataFormatter();
	    		 String val1= dataformat.formatCellValue(sheet.getRow(i).getCell(j));
	    	     System.out.println(val1);
                 */
	    	}
	    }
	    
	  }


	public static void main(String[] args) throws IOException 
	{
		// TODO Auto-generated method stub
		
		ExcelReadData data = new ExcelReadData();
		data.readExcel("E:\\Swati Study Material\\java programs\\ExcelSheetDemo\\StudentDetails.xlsx", "Sheet1");
		
		
	}

}
