package EXCEL.com;

import java.io.FileInputStream;

import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel_File {
	
	@SuppressWarnings("resource")
	public static void main(String[] args) throws IOException {
		FileInputStream file = new FileInputStream("./src/main/resources/demo_excel.ods");
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		Iterator<Row> riter = sheet.iterator();
		while (riter.hasNext()) {
			Row row = riter.next();
			Iterator<Cell> citer = row.cellIterator();
			while (citer.hasNext()) {
				Cell cell = citer.next();
				switch (cell.getCellType()) {
				case Cell.CELL_TYPE_NUMERIC:
					System.out.println(cell.getNumericCellValue() + "/t");
					break;
				case Cell.CELL_TYPE_STRING:
					System.out.println(cell.getStringCellValue());
					break;
				}
			}
			System.out.println("");
		}

		file.close();
	}
}
