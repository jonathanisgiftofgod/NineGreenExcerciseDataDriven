package greenExcercise;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class GTFKG {
	public static void main(String[] args) throws IOException {
		File f = new File("C:\\Users\\Britto\\eclipse-workspace\\GreenExcerciseDataDriven\\Excel\\Excercise.xlsx");
		FileInputStream str = new FileInputStream(f);
		Workbook w = new XSSFWorkbook(str);
		Sheet sheet = w.getSheet("StudentsData");
		Row row = sheet.getRow(0);
		Cell cell = row.getCell(0);
		String stringCellValue = cell.getStringCellValue();
		if(stringCellValue.equals("Hi")) {
			cell.setCellValue("Hellow");
		}
		FileOutputStream str1 = new FileOutputStream(f);
		w.write(str1);
	}	
}
