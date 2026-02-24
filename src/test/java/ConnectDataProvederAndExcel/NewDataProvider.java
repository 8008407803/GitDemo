package ConnectDataProvederAndExcel;

import java.io.FileInputStream;
import java.io.IOException;

import javax.sound.sampled.TargetDataLine;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class NewDataProvider {

    DataFormatter format = new DataFormatter();
	@Test(dataProvider = "driveTest")
	public void testCaseData(String relation, String occupation, String age) {
		System.out.println(relation);
		System.out.println(occupation);
		System.out.println(age);
	}
	
	@DataProvider(name="driveTest")
	public Object[][]  getData() throws IOException {
//		Object[][] data = {{"Amma","HouseWife",44}, {"Nanna","DailyWase",48}, {"Chelli","Accountant",22}};
//		return data;
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
				XSSFCell cell = row.getCell(j);
				data[i][j] = format.formatCellValue(cell);
				System.out.println(row.getCell(j));
			}
		}
		return data;
	}

}
