package pkg;

import java.io.IOException;

public class getcolcount {

	public static void main(String[] args) throws IOException {
		int y = XLUTILS.getcolumncount("C:\\demo\\Book1.xlsx", "EmpData",1);
		System.out.println(y);
	}

}
