package demo.poi.excel;

import java.util.ArrayList;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import poi.excel.ListSheets;
import poi.excel.ReadSheet;

/**
 * <p>
 *  ListSheetsDemo
 * </p>
 * A demonstration to use poi.excel.ListSheets
 * @author Chris Wong
 *
 */
public class ListSheetsDemo {
	
	private final static String txtFileName = "resource\\test.txt";
	private final static String xlsFileName = "resource\\test.xls";
	private final static String xlsxFileName = "resource\\test.xlsx";
	private final static String testxlsx1 = "resource\\ETL_DEV_Account.xlsx";
	
	/**
	 * Main program
	 * 
	 * @param args	arguments of running java
	 */
	public static void main(String[] args){
		// tests for listing out the names of sheets
		readExcel(txtFileName);
		readExcel(xlsFileName);
		readExcel(xlsxFileName);
		
		// tests for showing the cells of a single sheet in a xls file
		System.out.println("======================Tests for showing the cells of xls file==========================================");
		readExcel(txtFileName, 0);
		readExcel(xlsFileName, 0);
		readExcel(xlsFileName, 1);
		readExcel(xlsFileName, 2);

		
		// tests for showing the cells of a single sheet in a xlsx file
		System.out.println("======================Tests for showing the cells of xlsx file==========================================");
		readExcel(txtFileName, 0);
		readExcel(xlsxFileName, 0);
		readExcel(xlsxFileName, 1);
		readExcel(xlsxFileName, 2);
		
		//readExcel(testxlsx1, 0);
	}
	
	/**
	 * Read an excel file and list out the names of sheets
	 * 
	 * @param filename excel file name
	 */
	public static void readExcel(String filename){
		ArrayList<String> names = ListSheets.getNames(filename);
		if(names == null){
			System.out.println(filename + " is not an excel file");
		} else {
			for(String name: names){
				System.out.println(name);
			}
		}
	}
	

	/**
	 * Read an excel file and show rows of a single sheet
	 * 
	 * @param filename		excel file name
	 * @param sheetIndex	sheet index
	 */
	public static void readExcel(String filename, int sheetIndex){
		Object sheet = ListSheets.getSheet(filename, sheetIndex);
		if(sheet instanceof HSSFSheet){
			System.out.println("--------------------XLS file:" + filename + ", sheet " + sheetIndex + "------------------------");
		    ReadSheet.sheetIterate((HSSFSheet)sheet);
		} else if (sheet instanceof XSSFSheet){
			System.out.println("--------------------XLSX file:" + filename + ", sheet " + sheetIndex + "------------------------");
		    ReadSheet.sheetIterate((XSSFSheet)sheet);
		}
	}
}
