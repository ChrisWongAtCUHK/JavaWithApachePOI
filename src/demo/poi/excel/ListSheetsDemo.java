package demo.poi.excel;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Map.Entry;

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
	
	private final static String TXTFILENAME = "resource\\test.txt";
	private final static String XLSRILENAME = "resource\\test.xls";
	private final static String XLSXFILENAME = "resource\\test.xlsx";

	private final static String testxlsx1 = "resource\\ETL_DEV_Account.xlsx";
	
	/**
	 * Main program
	 * 
	 * @param args	arguments of running java
	 */
	public static void main(String[] args){
		
		// tests for listing out the names of sheets
		readExcel(TXTFILENAME);
		readExcel(XLSRILENAME);
		readExcel(XLSXFILENAME);
		
		// tests for showing the cells of a single sheet in a xls file
		System.out.println("======================Tests for showing the cells of xls file==========================================");
		readExcel(TXTFILENAME, 0);
		readExcel(XLSRILENAME, 0);
		readExcel(XLSRILENAME, 1);
		readExcel(XLSRILENAME, 2);

		
		// tests for showing the cells of a single sheet in a xlsx file
		System.out.println("======================Tests for showing the cells of xlsx file==========================================");
		readExcel(TXTFILENAME, 0);
		readExcel(XLSXFILENAME, 0);
		readExcel(XLSXFILENAME, 1);
		readExcel(XLSXFILENAME, 2);
		
		// tests for showing the cells of a single sheet in a xls file by getSheetObject2DArray
		System.out.println("======================Tests for showing the cells of xlsx file by getSheetObject2DArray ==========================================");
		readExcelObject2DArray(TXTFILENAME, 0);
		readExcelObject2DArray(XLSRILENAME, 0);
		readExcelObject2DArray(XLSRILENAME, 1);
		readExcelObject2DArray(XLSRILENAME, 2);
		
		// tests for showing all sheets of an excel file
		readAllSheets(TXTFILENAME);
		readAllSheets(XLSRILENAME);
		readAllSheets(XLSXFILENAME);
	}
	
	/**
	 * Read an excel file and list out the names of sheets
	 * 
	 * @param filename excel file name
	 */
	public static void readExcel(String filename){
		ArrayList<String> names = ListSheets.getNames(filename);
		if(names != null){
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
	
	/**
	 * Read an excel file and show rows of a single sheet by getSheetObject2DArray
	 * 
	 * @param filename		excel file name
	 * @param sheetIndex	sheet index
	 */
	public static void readExcelObject2DArray(String filename, int sheetIndex){
		ArrayList<ArrayList<Object>> Object2DArray = ListSheets.getSheetObject2DArray(filename, sheetIndex);
		
		if(Object2DArray == null){
			return;
		}
		
		for(ArrayList<Object> objects: Object2DArray){
			for(Object object: objects){
				System.out.print(object + "\t\t\t");
			}
			System.out.println();
		 }
	}
	
	/**
	 * Read an excel file and show rows of a single sheet by getSheetObject2DArray
	 * 
	 * @param filename		excel file name
	 * @param sheetIndex	sheet index
	 */
	public static void readExcelObject2DArray(String filename, String sheetName){
		ArrayList<ArrayList<Object>> Object2DArray = ListSheets.getSheetObject2DArray(filename, sheetName);
		
		if(Object2DArray == null){
			return;
		}
		
		for(ArrayList<Object> objects: Object2DArray){
			for(Object object: objects){
				System.out.print(object + "\t\t\t");
			}
			System.out.println();
		 }
	}

	/**
	 * Read an excel file and show all sheets
	 * 
	 * @param filename		excel file name
	 */
	public static void readAllSheets(String filename){
		HashMap<String, ArrayList<ArrayList<Object>>> sheets = ListSheets.getAllSheets(filename);
		
		// invalid excel file
		if(sheets == null){
			return;
		}
		
		Iterator<Entry<String, ArrayList<ArrayList<Object>>>> it = sheets.entrySet().iterator();
	    while (it.hasNext()) {
	        Map.Entry<String, ArrayList<ArrayList<Object>>> pairs = (Map.Entry<String, ArrayList<ArrayList<Object>>>)it.next();
	        
	        // key is the sheet name
	        System.out.println("+++++++++++++++++" + pairs.getKey() + "+++++++++++++++++");
	        ArrayList<ArrayList<Object>> sheetObject2DArray = pairs.getValue();
	        
	        // avoid invalid cases
	        if(sheetObject2DArray != null){
	        	// value is sheet content
		        for(ArrayList<Object> objects: sheetObject2DArray){
					for(Object object: objects){
						System.out.print(object + "\t\t\t");
					}
					System.out.println();
				}
	        }

	        it.remove(); // avoids a ConcurrentModificationException
	    }
	}
}
