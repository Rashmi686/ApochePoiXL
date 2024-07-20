package pkg;

import java.io.IOException;

public class getstringcellvalue {

	public static void main(String[] args) throws IOException {
		String data1 = XLUTILS.getstringdata("C:\\demo\\Book1.xlsx","EmpData", 3, 0);
		System.out.println(data1);
		

	}

}
