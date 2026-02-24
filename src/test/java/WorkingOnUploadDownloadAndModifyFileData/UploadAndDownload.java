package WorkingOnUploadDownloadAndModifyFileData;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;

public class UploadAndDownload {
	
	static String fileName = "C:\\Users\\HP\\Downloads\\download.xlsx";

	public static void main(String[] args) throws IOException {
		
		String fruitName = "Apple";
		String updatedValue = "600";
		
		WebDriver driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(7));
		driver.get("https://rahulshettyacademy.com/upload-download-test/index.html");
		
		// Download
		driver.findElement(By.cssSelector("#downloadButton")).click();
		
		// Edit Excel
		int col = getColumnNumber(fileName, "price");
		int row = getRowNumber(fileName, fruitName);
		Assert.assertTrue(updateCell(fileName,row,col,updatedValue));
		
		
		// Upload file
		WebElement upload = driver.findElement(By.cssSelector("input[type='file']"));
		upload.sendKeys(fileName);
		
		// Wait for success message to show up and wait for disappear
		WebElement successMsgApr = driver.findElement(By.className("Toastify__toast-body"));
		WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
		String eleApr = wait.until(ExpectedConditions.visibilityOf(successMsgApr)).getText();
		if(eleApr.contains("Success")) {
			System.out.println("Success message is displayed successfully");
		}
		// check for message is disappeared
		boolean eleDispr = wait.until(ExpectedConditions.invisibilityOf(successMsgApr));
		if(eleDispr) {
			System.out.println("Success message is disappeared successfully");
		}
		
		// Verify updated excel data showing in the web table
		String priceClmn = driver.findElement(By.xpath("//div[text()='Price']")).getDomAttribute("data-column-id");
		WebElement eleXpath = driver.findElement(By.xpath("//div[text()='"+fruitName+"']/parent::div/parent::div/div["+priceClmn+"]/div"));
		// Price of an 'Apple'
		String price = eleXpath.getText();
		Assert.assertEquals("345", price);
		System.out.println(price);	

	}
	
	private static int getRowNumber(String fileName, String textName) throws IOException {
		ArrayList<String> a = new ArrayList<String>();
		// C:\\Users\\ADMIN\\Desktop\\geeks.xlsx
		FileInputStream file = new FileInputStream(fileName);
		XSSFWorkbook workbook = new XSSFWorkbook(file);

		XSSFSheet sheet = workbook.getSheet("Sheet1");
		Iterator<Row> rIterator = sheet.iterator();
		int k = 1;
		int rowIndex = -1;
		int coloumn = 0;
		
		while(rIterator.hasNext()) {
			Row row = rIterator.next();
			Iterator<org.apache.poi.ss.usermodel.Cell> rit = row.cellIterator();
			while(rit.hasNext()) {
				org.apache.poi.ss.usermodel.Cell cell = rit.next();
				if (cell.getCellType() == CellType.STRING && cell.getStringCellValue().equalsIgnoreCase(textName)) {
					rowIndex = k;
				}
			}
			k++;
		}
		return rowIndex;
	}
	
	private static int getColumnNumber(String fileName, String columnName) throws IOException {
		int coloumn = 0;
		ArrayList<String> a = new ArrayList<String>();
		// C:\\Users\\ADMIN\\Desktop\\geeks.xlsx
		FileInputStream file = new FileInputStream(fileName);
		XSSFWorkbook workbook = new XSSFWorkbook(file);

		XSSFSheet sheet = workbook.getSheet("Sheet1");

		Iterator<Row> rIterator = sheet.iterator();
		Row row = rIterator.next();
		Iterator<org.apache.poi.ss.usermodel.Cell> rit = row.cellIterator();

		int k = 1;

		while (rit.hasNext()) {
			org.apache.poi.ss.usermodel.Cell cell = rit.next();

			if (cell.getStringCellValue().equalsIgnoreCase(columnName)) {
				coloumn = k;
			}
			k++;
		}
		System.out.println(coloumn);

		return coloumn;
	}
	
	private static boolean updateCell(String fileName, int row, int col, String updatedValue) throws IOException {
		ArrayList<String> a = new ArrayList<String>();
		// C:\\Users\\ADMIN\\Desktop\\geeks.xlsx
		FileInputStream file = new FileInputStream(fileName);
		XSSFWorkbook workbook = new XSSFWorkbook(file);

		XSSFSheet sheet = workbook.getSheet("Sheet1");
		Row rowField = sheet.getRow(row-1);
		Cell cellField = rowField.getCell(col-1);
		cellField.setCellValue(updatedValue);
		FileOutputStream fos = new FileOutputStream(fileName);
		workbook.write(fos);
		workbook.close();
		file.close();
		
		return true;
	}
}
