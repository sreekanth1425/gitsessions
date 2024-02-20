package sample;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelTryCatchNestedForLoop {

	public static void main(String[] args) {
		try {
            // Specify the path to your Excel file
			// Specify the path to your Excel file
            String excelFilePath = "C:\\1\\DeleteAfterPractice\\src\\deletePractice\\venkatesh.xlsx";

            // Create a FileInputStream to read the Excel file
            FileInputStream fis = new FileInputStream(new File(excelFilePath));

            // Create an XSSFWorkbook object representing the Excel workbook
            XSSFWorkbook workbook = new XSSFWorkbook(fis);

            // Get the first sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0);

            // Iterate through each row in the sheet
            for (Row row : sheet) {
                // Iterate through each cell in the row
                for (Cell cell : row) {
                    // Print the cell value
                    System.out.print(cell.toString() + "\t");
                }
                System.out.println(); // Move to the next line after each row
            }

            // Close the workbook and file input stream
            workbook.close();
            fis.close();

        } catch (IOException e) {
            e.printStackTrace();
        }
	}

}
