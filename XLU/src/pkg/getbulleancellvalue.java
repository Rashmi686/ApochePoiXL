package pkg;

import java.io.IOException;

public class getbulleancellvalue {
public static void main(String[] args) throws IOException {
	Boolean val = XLUTILS.getbulleancellvalue("C:\\demo\\Book1.xlsx", "EmpData", 1, 3);
	System.out.println(val);
}
}
