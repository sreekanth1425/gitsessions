package sample;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel {

	public static void main(String[] args) throws IOException {
		FileInputStream fis = new FileInputStream("C:\\1\\DeleteAfterPractice\\src\\deletePractice\\venkatesh.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheet("Sheet1");
		System.out.println(sheet.getSheetName());
		System.out.println(sheet.getLastRowNum());
		System.out.println("Before updating Cell Data: " + sheet.getRow(4).getCell(1));
		// Write the data to excel file
		XSSFCell cell = sheet.getRow(2).getCell(1);
		cell.setCellValue("Mahesh");
		fis.close();

		FileOutputStream fileOut = new FileOutputStream(
				"C:\\1\\DeleteAfterPractice\\src\\deletePractice\\venkatesh.xlsx");
		workbook.write(fileOut);
		System.out.println("Updated data after write is done :" + cell.getStringCellValue());
		fileOut.close();

	}

}
