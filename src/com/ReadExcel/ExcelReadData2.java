package com.ReadExcel;

import java.io.FileInputStream;

import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.DataFormatter;

public class ExcelReadData2 
{
	public void readExcel(String filename,String sheetname) throws IOException
	{
		int arrayexcel[][]=null;
		FileInputStream fis = new FileInputStream(filename);
		
		HSSFWorkbook wb = new HSSFWorkbook(fis);
		HSSFSheet sheet = wb.getSheet(sheetname);
		HSSFRow row = sheet.getRow(2);
		HSSFCell cell = row.getCell(2);
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
		
		ExcelReadData2 data = new ExcelReadData2();
		data.readExcel("E:\\Swati Study Material\\java programs\\ExcelSheetDemo\\studentDetail1.xls", "Sheet1");
		
		
	}

}
