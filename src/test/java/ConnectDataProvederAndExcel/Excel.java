package ConnectDataProvederAndExcel;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class Excel {
	
	@Test
	public void getExcel() throws IOException {
		FileInputStream file = new FileInputStream("C:\\Users\\HP\\Downloads\\EcxelDriven.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheetAt(0);
		int rows = sheet.getPhysicalNumberOfRows();
		XSSFRow row = sheet.getRow(0);
		int cellCount = row.getLastCellNum();
		
		Object data[][] = new Object[rows-1][cellCount];
		
		for (int i=0;i<rows-1;i++) {
			row = sheet.getRow(i+1);
			for (int j=0;j<cellCount;j++) {
				data[i][j] = row.getCell(j);
				System.out.println(row.getCell(j));
			}
		}
	}

}
