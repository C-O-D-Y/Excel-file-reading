package ReadFile;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/*
 * In this Class i'm reading excel file using maven
 */
public class FetchFrom_xlsFile {
	/*
	 * In this method i'm reading excel file using dependency of apache-poi
	 */
	public static void main(String[] args) throws IOException {
		FileInputStream fis = new FileInputStream(new File("C:\\Users\\saurabh.chauhan\\Desktop\\Book.xlsx"));
		// Class used to read excel file and read the data
		XSSFWorkbook file = new XSSFWorkbook(fis);
		XSSFSheet worksheet = file.getSheetAt(0);
		// iterating through rows and getting row number
		Iterator<Row> rows = worksheet.iterator();
		rows.next();
		// in while loop, iterating with column to get the values through the cell
		while (rows.hasNext()) {
			Row row = rows.next();
			Iterator<Cell> iterator = row.cellIterator();
			while (iterator.hasNext()) {
				Cell cell = iterator.next();
				System.out.print(cell.toString() + " ");
			}
			System.out.println();
		}
		fis.close();
	}

}
