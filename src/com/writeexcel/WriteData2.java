package com.writeexcel;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteData2 
{
	public void setCellData(String filenm,String sheetnm,int rownum,int colnum,String dataval) throws IOException
	{
		FileInputStream fis1 = new FileInputStream(filenm);

		HSSFWorkbook wb1 = new HSSFWorkbook(fis1);
		HSSFSheet sheet = wb1.getSheet(sheetnm);
		HSSFRow row = sheet.createRow(rownum);
		HSSFCell cell = row.createCell(colnum);	
		cell.setCellValue(dataval);

		FileOutputStream fio = new FileOutputStream(filenm);
		wb1.write(fio);

	}
}