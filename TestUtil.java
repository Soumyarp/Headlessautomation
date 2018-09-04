package Utility;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.reporters.jq.Main;

public class TestUtil {

	public static String TESTDATA_SHEET_PATH = System.getProperty("user.dir") + "/Testdata/Data.xlsx";
	static Workbook book;
	static Sheet sheet;

	public static Object[][] getTestData(String sheetName) {
		// System.out.println("hiiiiiiiiiiiiii");
		FileInputStream file = null;
		try {
			file = new FileInputStream(TESTDATA_SHEET_PATH);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		try {
			book = WorkbookFactory.create(file);
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		sheet = book.getSheet(sheetName);

		Object[][] data = new Object[sheet.getLastRowNum()][sheet.getRow(0).getLastCellNum()];
		for (int i = 0; i < sheet.getLastRowNum(); i++) {
			for (int j = 0; j < sheet.getRow(0).getLastCellNum(); j++) {
				data[i][j] = sheet.getRow(i + 1).getCell(j).toString();

			}

		}
		return data;

	}
	
	public void updateExcel(String sheetName, String Username, String password, String Status) throws IOException{
		FileInputStream file = new FileInputStream(new File(TESTDATA_SHEET_PATH));
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet=workbook.getSheet(sheetName);
		int totalrow=sheet.getLastRowNum()+1;
		System.out.println(totalrow);
		for (int i = 0; i < totalrow; i++) {
			XSSFRow row=sheet.getRow(i);
		
			String cell=row.getCell(1).getStringCellValue();
		//	System.out.println(cell);
			if(cell.contains(password)){
			
		       
				row.createCell(2).setCellValue(Status);
				file.close();
				
				System.out.println("Column added successfully");
				FileOutputStream fo = new FileOutputStream(new File(TESTDATA_SHEET_PATH));
				workbook.write(fo);
				fo.close();
				break;
				
			}
			
			
			
		}
		
	 
		
	}
	public static void main(String[] args) throws IOException {
		String sheetName = "Logindata";
		TestUtil util = new TestUtil();
		util.updateExcel(sheetName, "Test1", "Test", "Pass");
		
	}
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
}
