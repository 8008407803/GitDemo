import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import com.google.common.collect.Table.Cell;

public class DataDriven {
	
	public ArrayList<String> getData(String TestCase) throws IOException {
		
		ArrayList<String> a = new ArrayList<String>();
		// C:\\Users\\ADMIN\\Desktop\\geeks.xlsx
		FileInputStream file = new FileInputStream("C:\\Users\\HP\\Downloads\\TestData.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		
		int sheets = workbook.getNumberOfSheets();
		
		for (int i=0;i<sheets;i++) {
			if(workbook.getSheetName(i).equalsIgnoreCase("MainSheet1")) {
				XSSFSheet sheet = workbook.getSheetAt(i);
				
				// Identify the 'ID' column by scannin the entire 1st row
				Iterator<Row> rIterator = sheet.iterator();
				Row row = rIterator.next();
				Iterator<org.apache.poi.ss.usermodel.Cell> rit = row.cellIterator();
				
				int k = 0;
				int coloumn = 0;
				
				while (rit.hasNext()) {
					org.apache.poi.ss.usermodel.Cell cell = rit.next();
					 
					if (cell.getStringCellValue().equalsIgnoreCase("ID")) {
					coloumn = k;
					}
					k++;
				}
				System.out.println(coloumn);
				
				// Once coloumn is identified then scan entire testcase coloumn to identify specific testcase row
				while(rIterator.hasNext()) {
					Row r = rIterator.next();
					if(r.getCell(coloumn).getStringCellValue().equalsIgnoreCase(TestCase)) {
						// After you grab 3rd row  = pull all the data of that row and feed into test
						Iterator<org.apache.poi.ss.usermodel.Cell> cv = r.cellIterator();
						while(cv.hasNext()) {
							org.apache.poi.ss.usermodel.Cell c = cv.next();
							if(c.getCellType()==CellType.STRING) {
							a.add(c.getStringCellValue());
							}
							else {
								a.add(NumberToTextConverter.toText(c.getNumericCellValue()));
							}
						}
					}
				}
			}
		}
		return a;
		
	}

	public static void main(String[] args) {
		
	}

}
