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
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class prac3 {
	public static void main(String[] args) throws IOException {
		FileInputStream fi=new FileInputStream("c:\\demo\\test23.xlsx");
		XSSFWorkbook wb=new XSSFWorkbook(fi);
		Sheet ws = wb.getSheet("Data");
		Row row = ws.getRow(2);
		Cell cell = row.createCell(0);
		cell.setCellValue("PASS");
		CellStyle style = wb.createCellStyle();
		style.setFillForegroundColor(IndexedColors.RED.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		cell.setCellStyle(style);
		FileOutputStream fo=new FileOutputStream("c:\\demo\\test234.xlsx");
		wb.write(fo);
	}
}