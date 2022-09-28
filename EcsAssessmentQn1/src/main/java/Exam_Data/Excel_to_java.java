package Exam_Data;

import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel_to_java {
	public String CellData(int row, int col, String path) throws IOException {

		XSSFWorkbook Workbook = new XSSFWorkbook(path);
		XSSFSheet sheet = Workbook.getSheet("Sheet1");
		try {
			return sheet.getRow(row).getCell(col).getStringCellValue();
		} catch (RuntimeException e) {
			return "";
		} finally {
			Workbook.close();
		}
	}

	public int numCellData(int row, int col, String path) throws IOException {

		XSSFWorkbook Workbook = new XSSFWorkbook(path);
		XSSFSheet sheet = Workbook.getSheet("sheet1");
		try {
			return (int) sheet.getRow(row).getCell(col).getNumericCellValue();
		} catch (RuntimeException e) {

			return 0;
		} finally {
			Workbook.close();
		}

	}

	public int rowCount(String path) throws IOException {

		XSSFWorkbook workbook = new XSSFWorkbook(path);
		XSSFSheet sheet = workbook.getSheet("sheet1");
		int rows = sheet.getPhysicalNumberOfRows();
		workbook.close();
		return rows;
		 
				
			
	}

	public String gradeCalculation(int Score) {

		String Grade = "";
		if (Score > 90) {
			Grade = "A1";
		}
		if (Score <= 90) {
			Grade = "A2";
		}
		if (Score <= 80) {
			Grade = "B1";
		}
		if (Score <= 70) {
			Grade = "B2";
		}
		if (Score <= 60) {
			Grade = "C1";
		}
		if (Score <= 50) {
			Grade = "C2";
		}
		if (Score <= 40) {
			Grade = "D";
		}
		if (Score <= 32) {
			Grade = "E1";
		}
		if (Score <= 20) {
			Grade = "E2";
		}

		return Grade;
	}

	public float gradePointCalculation(int Score) {

		int GradePoint = 0;

		if (Score > 90) {
			GradePoint = 10;
		}
		if (Score <= 90) {
			GradePoint = 9;
		}
		if (Score <= 80) {
			GradePoint = 8;
		}
		if (Score <= 70) {
			GradePoint = 7;
		}
		if (Score <= 60) {
			GradePoint = 6;
		}
		if (Score <= 50) {
			GradePoint = 5;
		}
		if (Score <= 40) {
			GradePoint = 4;
		}
		return GradePoint;
	}

}
