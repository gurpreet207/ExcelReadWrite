package Readexcel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.usermodel.Cell;



public class reading {
	public static String TESTDATA_SHEET_PATH = System.getProperty("user.dir")+"\\InputData - Copy.xlsx";
	static Workbook book;
	static Sheet sheet;
	static Row row;
	static Cell Cell;
	static String sheetName = "info";
	
	public static Object[][] getTestData(String sheetName) {
		FileOutputStream fout=null;
		FileInputStream file = null;
		try {
			file = new FileInputStream(TESTDATA_SHEET_PATH);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		try {
			book = WorkbookFactory.create(file);//Creates the appropriate HSSFWorkbook / XSSFWorkbook from the given File, which must exist and be readable.
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		sheet = book.getSheet(sheetName);
		Object[][] data = new Object[sheet.getLastRowNum()][sheet.getRow(0).getLastCellNum()];
		System.out.println(sheet.getLastRowNum() + "--------" +		sheet.getRow(0).getLastCellNum());
		for (int i = 0; i < sheet.getLastRowNum(); i++) {
			for (int k = 0; k < sheet.getRow(0).getLastCellNum(); k++) {
				data[i][k] = sheet.getRow(i + 1).getCell(k).toString();
				row=sheet.getRow(i+1);
				Cell = row.getCell(k);
				 //System.out.print(data[i][k] + " ");
				System.out.print(" "+ Cell +" ");
			}
		}
		
		//sheet.getRow(0).createCell(3).setCellValue("Pass");
		
		for (int i = 0; i < sheet.getLastRowNum(); i++) {
			for (int k = 0; k < sheet.getRow(0).getLastCellNum(); k++) {
				sheet.getRow(i+1).createCell(3).setCellValue("Pass");
			}
		}
		
		try {
			fout =new FileOutputStream(TESTDATA_SHEET_PATH);
		} catch (FileNotFoundException e) {
			
			e.printStackTrace();
		}
		
		try {
			book.write(fout);
			book.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		
		return data;
	}
	
		
	public static void main(String[] args) {
	Object data1[][] = getTestData(sheetName);
	
        //System.out.println(data1[0][3]);
	}

}