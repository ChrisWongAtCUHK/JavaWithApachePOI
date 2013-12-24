package demo.poi.excel;

import java.util.ArrayList;

import poi.excel.ListSheets;

/**
 * <p>
 *  ListSheetsDemo
 * </p>
 * A demonstration to use poi.excel.ListSheets
 * @author Chris Wong
 *
 */
public class ListSheetsDemo {
	
	static String txtFileName = "resource\\test.txt";
	static String xlsFileName = "resource\\test.xls";
	static String xlsxFileName = "resource\\test.xlsx";
	static String testxlsx1 = "resource\\ETL_DEV_Account.xlsx";
	
	public static void main(String[] args){
		readExcel(txtFileName);
		readExcel(xlsFileName);
		readExcel(xlsxFileName);
	}
	
	/**
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
}
