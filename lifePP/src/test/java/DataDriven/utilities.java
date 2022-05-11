package DataDriven;

import java.io.IOException;

public class utilities {

	public static void main(String[] args) {
	
	}
	private static void GetData() throws IOException {
		// TODO Auto-generated method stub
		ExcelDriven getData = new ExcelDriven();
		System.out.println(getData.ExcelGetData().get(1));
	}

}
	

