package base;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;

import io.github.bonigarcia.wdm.WebDriverManager;

public class BaseClass {
	public static WebDriver driver;
	public void launchUrl(String url) {
		WebDriverManager.chromedriver().setup();
		driver = new ChromeDriver();
		driver.get(url);
		driver.manage().window().maximize();
	}
	public void enterText(WebElement id, String text) {
		id.sendKeys(text);
	}
	public void btnClick(WebElement id) {
		id.click();
	}
	public void Clear(WebElement id) {
		id.clear();
	}
	public void selectById(WebElement id, int index) {
		Select s = new Select(id);
		s.selectByIndex(index);
	}
	public void selectByValue(WebElement id, String value) {
		Select s = new Select(id);
		s.selectByValue(value);
	}
	public void selectByVisibleText(WebElement id, String value) {
		Select s = new Select(id);
		s.selectByValue(value);
	}
	public void startTime() {
		Date d = new Date();
		System.out.println("Starting Time "+d);
	}
	public void endTime() {
		Date d = new Date();
		System.out.println("Ending Time "+d);
	}
	
	public String inputValues(String Locator) {
		WebElement element = driver.findElement(By.id(Locator));
		String input = element.getAttribute("value");
		System.out.println(input);
		return input;
	}
	public String readExcel(String sheet, int row, int column) throws IOException {
		File f = new File("C:\\Users\\Britto\\eclipse-workspace\\GreenExcerciseDataDriven\\Excel\\Naukri.xlsx");
		FileInputStream fi = new FileInputStream(f);
		Workbook w = new XSSFWorkbook(fi);
		Sheet s = w.getSheet(sheet);
		Row r = s.getRow(row);
		Cell c = r.getCell(column);
		int cellType = c.getCellType();
		String value = null;
		if (cellType==1) {
			value = c.getStringCellValue();
		} else if(DateUtil.isCellDateFormatted(c)){
			Date d = c.getDateCellValue();
			SimpleDateFormat sf = new SimpleDateFormat("dd/MM/yyyy"); 
			value = sf.format(d);
		}
		else {
			double d = c.getNumericCellValue();
			long l = (long)d;
			value = String.valueOf(l);
		}
		return value;
	}
	public void writeExcel(int row, int column, String order) throws IOException {
		File f = new File("C:\\Users\\Britto\\eclipse-workspace\\GreenExcerciseDataDriven\\Excel\\Output.xlsx");
		Workbook w = new XSSFWorkbook();
		Sheet s = w.createSheet("Sheet1");
		Row r = s.createRow(row);
		Cell c = r.createCell(column);
		c.setCellValue(order);
		FileOutputStream fo = new FileOutputStream(f);
		w.write(fo);
		fo.close();
	}
}
