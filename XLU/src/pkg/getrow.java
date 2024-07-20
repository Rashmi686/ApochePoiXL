package pkg;

import java.io.IOException;

public class getrow {
 public static void main(String[] args) throws IOException {
	int x = XLUTILS.getrowcount("C:\\demo\\Book1.xlsx", "EmpData");
	System.out.println(x);
}
}
