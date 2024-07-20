package pkg;

import java.io.IOException;

public class cellstyle {

	public static void main(String[] args) throws IOException {
		XLUTILS.setdata("C:\\demo\\Book1.xlsx", "EmpData",1, 3, "PASS");
		XLUTILS.fillforegroundcolour("C:\\demo\\Book1.xlsx", "EmpData",1, 3, "PASS");
	}
}
