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

public class practice {

	public static void main(String[] args) throws IOException {
		FileInputStream fi = new FileInputStream("C:\\demo\\Book1.xlsx");
		Workbook wb = new XSSFWorkbook(fi);
		Sheet ws = wb.getSheet("EmpData");
		Row row = ws.getRow(1);
	Cell cell = row.createCell(4);
	cell.setCellValue("PASS");
	
	CellStyle style = wb.createCellStyle();
	style.setFillForegroundColor(IndexedColors.BRIGHT_GREEN.getIndex());
	style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	cell.setCellStyle(style);
	FileOutputStream fo = new FileOutputStream("C:\\demo\\Book14.xlsx");
	wb.write(fo);
		
	}

}
