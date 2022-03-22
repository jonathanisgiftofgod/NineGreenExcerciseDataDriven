package greenExcercise;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import org.openqa.selenium.By;

import base.BaseClass;

public class GreenExcerciseDataDriven {
	public static void main(String[] args) throws IOException {
		//File Location
		File f = new File("C:\\Users\\Britto\\eclipse-workspace\\GreenExcerciseDataDriven\\Excel\\Excercise.xlsx");
		Workbook w = new XSSFWorkbook();
		Sheet s = w.createSheet("StudentsData");
		int cellCount=0;
		for (int i = 0; i <=9; i++) {
			Row r = s.createRow(i);
			for (int j = 0; j < 3; j++) {
				System.out.println("Enter the value : ");
				Cell c = r.createCell(j);
				Scanner sc = new Scanner(System.in);
				String value = sc.nextLine(); 
				c.setCellValue(value);
				cellCount++;
			}
		}
		int physicalNumberOfRows = s.getPhysicalNumberOfRows();
		System.out.println("Rows Count : "+physicalNumberOfRows);
		System.out.println("Cell count " + cellCount);
		FileOutputStream fo = new FileOutputStream(f);
		w.write(fo);
		fo.close();
	}
}
