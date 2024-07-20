package pkg;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class prac2 {
	public static void main(String[] args) throws IOException {
		FileInputStream fi=new FileInputStream("c:\\demo\\test23.xlsx");
		XSSFWorkbook wb=new XSSFWorkbook(fi);
		Sheet sh=wb.getSheet("Data");
		Row row=sh.getRow(1);
		Cell cell =row.getCell(0);
		
		
}
}