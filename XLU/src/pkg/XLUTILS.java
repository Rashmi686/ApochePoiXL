package pkg;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;






public class XLUTILS 
{
	
	
	
public static FileInputStream fi;
public static FileOutputStream fo;
public static Workbook wb;
public static Sheet ws;
public static Row row;
public static Cell cell;
public static CellStyle style;

	public static int getrowcount(String xlfile,String xlsheet) throws IOException
	{
		fi = new FileInputStream(xlfile);
		wb = new XSSFWorkbook(fi);
		ws = wb.getSheet(xlsheet);
		int rowcount = ws.getLastRowNum();
		wb.close();
		return rowcount;	
	}

	public static int getcolumncount(String xlfile,String xlsheet, int rownum) throws IOException
	{
		fi = new FileInputStream(xlfile);
		wb = new XSSFWorkbook(fi);
		ws = wb.getSheet(xlsheet);
		row = ws.getRow(rownum);
		Short colcount = row.getLastCellNum();
		wb.close();
	
	return colcount;
	
}
	public static String getstringdata(String xlfile,String xlsheet, int rownum,int colnum) throws IOException
	{
		fi = new FileInputStream(xlfile);
		wb = new XSSFWorkbook(fi);
		ws = wb.getSheet(xlsheet);
		row = ws.getRow(rownum);
	cell = row.getCell(colnum);
	String data = cell.getStringCellValue();
	wb.close();
	return data;
	
	}
	public static Boolean getbulleancellvalue(String xlfile,String xlsheet, int rownum,int colnum) throws IOException
	{
		FileInputStream fi = new FileInputStream(xlfile);
		wb = new XSSFWorkbook(fi);
	  ws = wb.getSheet(xlsheet);
	  row =ws.getRow(colnum);
	 cell = row.getCell(colnum);
	 Boolean res = cell.getBooleanCellValue();
	 wb.close();
	 return res;
	
		
	}
	
	public static void setdata(String xlfile,String xlsheet, int rownum,int colnum,String data) throws IOException
	{
		FileInputStream fi = new FileInputStream(xlfile);
		wb = new XSSFWorkbook(fi);
		ws = wb.getSheet(xlsheet);
		row = ws.getRow(rownum);
	cell = row.createCell(colnum);
	cell.setCellValue(data);
 fo = new FileOutputStream(xlfile);
 wb.write(fo);
 wb.close();

	}
	public static void fillforegroundcolour(String xlfile,String xlsheet, int rownum,int colnum,String data) throws IOException
	{
		FileInputStream fi = new FileInputStream(xlfile);
		wb = new XSSFWorkbook(fi);
		ws = wb.getSheet(xlsheet);
		row = ws.getRow(rownum);
		cell = row.getCell(colnum);
		style = wb.createCellStyle();
		
		style.setFillForegroundColor(IndexedColors.RED.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		cell.setCellStyle(style);
		fo = new FileOutputStream(xlfile);
		 wb.write(fo);
		 wb.close();
	}
}
		



	

