package com.writeexcel;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteData 
{
	public void setCellData(String filenm,String sheetnm,int rownum,int colnum,String dataval) throws IOException
	{
		FileInputStream fis1 = new FileInputStream(filenm);

		XSSFWorkbook wb1 = new XSSFWorkbook(fis1);
		XSSFSheet sheet = wb1.getSheet(sheetnm);
		XSSFRow row = sheet.createRow(rownum);
		XSSFCell cell = row.createCell(colnum);	
		cell.setCellValue(dataval);

		FileOutputStream fio = new FileOutputStream(filenm);
		wb1.write(fio);

	}
}