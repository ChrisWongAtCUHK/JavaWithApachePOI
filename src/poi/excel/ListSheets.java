package poi.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;

//For .xls, libraries included: poi-3.9-20121203.jar
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

//For .xlsx, so many additional libraries must included: ooxml-lib/dom4j-1.6.1.jar, ooxml-lib/xmlbeans-2.3.0.jar, poi-ooxml-3.9-20121203.jar, poi-ooxml-schemas-3.9-20121203.jar
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import static java.lang.System.out;

/**
 * <p>
 *  ListSheets
 * </p>
 * List the contents of sheets in a single xls/xlsx
 * @author Chris Wong
 *
 */
public class ListSheets {
	private final static String INVALIDSHEETINDEX = "INVALID_SHEET_INDEX";
	private final static String INVALIDSHEETNAME = "INVALID_SHEET_NAME";
	private final static String NONEXCEL = "NON_EXCEL";
	final static String XLSFILENAME = "resource\\sheets.xls";
	final static String XLSXFILENAME = "resource\\sheets.xlsx";

	/**
	 * @param filename	filename name of excel file
	 * @return			number of sheets in the excel file	
	 */
	public static int getNumberOfSheets(String filename){
		// Check the type of excel file, use HSSF or XSSF accordingly
		String excelType = ExcelFile.excelType(filename);
		
		if(excelType == null){
			// neither xls nor xlsx
			System.err.println(filename + " is not excel.");
			return 0;
		} else if(excelType.equals("xls")){
			 try {
				// get the file input stream
				FileInputStream file = new FileInputStream(new File(filename));
				
				// get the workbook instance for XLS file 
			    HSSFWorkbook workbook = new HSSFWorkbook(file);
			    
			    // get the size of workbook
			    return workbook.getNumberOfSheets();
			    
			} catch (FileNotFoundException e) {
			    e.printStackTrace();
			} catch (IOException e) {
			    e.printStackTrace();
			}
		} else if(excelType.equals("xlsx")){
			 try {
				// get the file input stream
				FileInputStream xfile = new FileInputStream(new File(filename));
				
				// get the workbook instance for XLSX file 
				 XSSFWorkbook xworkbook = new XSSFWorkbook(xfile);
			    
			    // get the size of workbook
				 return xworkbook.getNumberOfSheets();
			    
			} catch (FileNotFoundException e) {
			    e.printStackTrace();
			} catch (IOException e) {
			    e.printStackTrace();
			}
		}
				
		return 0;
	}

	/**
	 * Get names of sheets
	 * 
	 * @param 		filename name of excel file
	 * @return 		names of sheets
	 */
	public static ArrayList<String> getNames(String filename){
		// Check the type of excel file, use HSSF or XSSF accordingly
		String excelType = ExcelFile.excelType(filename);
		
		// ArrayList to store names of sheets
		ArrayList<String> names = new ArrayList<String>();
		
		if(excelType == null){
			// neither xls nor xlsx
			System.err.println(filename + " is not excel.");
			return null;
		} else if(excelType.equals("xls")){
			 try {
				// get the file input stream
				FileInputStream file = new FileInputStream(new File(filename));
				
				// get the workbook instance for XLS file 
			    HSSFWorkbook workbook = new HSSFWorkbook(file);
			    
			    // get the size of workbook
			    int numberOfSheets = workbook.getNumberOfSheets();
			    
			    // loop through the workbook, add the name of sheet to the arraylist one by one
			    for(int i = 0; i < numberOfSheets; i++){
			    	names.add(workbook.getSheetName(i));
			    }
			    
			} catch (FileNotFoundException e) {
			    e.printStackTrace();
			} catch (IOException e) {
			    e.printStackTrace();
			} catch (IllegalArgumentException e){
				System.err.println(filename + " is not excel.");
				return null;
			}
		} else if(excelType.equals("xlsx")){
			 try {
				// get the file input stream
				FileInputStream xfile = new FileInputStream(new File(filename));
				
				// get the workbook instance for XLSX file 
				 XSSFWorkbook xworkbook = new XSSFWorkbook(xfile);
			    
			    // get the size of workbook
			    int numberOfSheets = xworkbook.getNumberOfSheets();
	
			    // loop through the workbook, add the name of sheet to the arraylist one by one
			    for(int i = 0; i < numberOfSheets; i++){
			    	names.add(xworkbook.getSheetName(i));
			    }
			    
			} catch (FileNotFoundException e) {
			    e.printStackTrace();
			} catch (IOException e) {
			    e.printStackTrace();
			} catch (IllegalArgumentException e){
				System.err.println(filename + " is not excel.");
				return null;
			}
		}
		
		return names;
	}

	/**
	 * Get the rows of a single sheet, store the content in the form of string
	 * 
	 * @param filename		filename name of excel file
	 * @param sheetIndex	sheet index
	 * @return				rows of a single
	 */
	public static Object getSheet(String filename, int sheetIndex){
		// Check the type of excel file, use HSSF or XSSF accordingly
		String excelType = ExcelFile.excelType(filename);

		if(excelType == null){
			// neither xls nor xlsx
			
			return null;
		} else if(excelType.equals("xls")){

			try {
				// get the file input stream
				FileInputStream file = new FileInputStream(new File(filename));
				
				// get the workbook instance for XLS file 
			    HSSFWorkbook workbook = new HSSFWorkbook(file);
			    
			    // get a single sheet in the workbook
			    HSSFSheet sheet = workbook.getSheetAt(sheetIndex);
			    
			    return sheet;
			    
			    
			} catch (FileNotFoundException e) {
			    e.printStackTrace();
			} catch (IOException e) {
			    e.printStackTrace();
			}
		} else if(excelType.equals("xlsx")){
			try {
				// get the file input stream
				FileInputStream xfile = new FileInputStream(new File(filename));
				
				// get the workbook instance for XLSX file 
				 XSSFWorkbook xworkbook = new XSSFWorkbook(xfile);
			    
				// get a single sheet in the workbook
				XSSFSheet sheet = xworkbook.getSheetAt(sheetIndex);
	
				return sheet;
			    
			} catch (FileNotFoundException e) {
			    e.printStackTrace();
			} catch (IOException e) {
			    e.printStackTrace();
			} catch (IllegalArgumentException e){
				return INVALIDSHEETINDEX;
			}
		}
		
		return null;
	}
	
	/**
	 * Get the rows of a single sheet, store the content in the form of string
	 * 
	 * @param filename		file name of excel file
	 * @param sheetName		sheet name
	 * @return				rows of a single
	 */
	public static Object getSheet(String filename, String sheetName){
		// Check the type of excel file, use HSSF or XSSF accordingly
		String excelType = ExcelFile.excelType(filename);

		if(excelType == null){
			// neither xls nor xlsx
			return null;
		} else if(excelType.equals("xls")){

			try {
				// get the file input stream
				FileInputStream file = new FileInputStream(new File(filename));
				
				// get the workbook instance for XLS file 
			    HSSFWorkbook workbook = new HSSFWorkbook(file);
			    
			    // get a single sheet in the workbook
			    HSSFSheet sheet = workbook.getSheetAt(workbook.getSheetIndex(sheetName));
			    
			    return sheet;
			    
			    
			} catch (FileNotFoundException e) {
			    e.printStackTrace();
			} catch (IOException e) {
			    e.printStackTrace();
			}
		} else if(excelType.equals("xlsx")){
			try {
				// get the file input stream
				FileInputStream xfile = new FileInputStream(new File(filename));
				
				// get the workbook instance for XLSX file 
				 XSSFWorkbook xworkbook = new XSSFWorkbook(xfile);
			    
				// get a single sheet in the workbook
				XSSFSheet sheet = xworkbook.getSheetAt(xworkbook.getSheetIndex(sheetName));
	
				return sheet;
			    
			} catch (FileNotFoundException e) {
			    e.printStackTrace();
			} catch (IOException e) {
			    e.printStackTrace();
			} catch (IllegalArgumentException e){
				return INVALIDSHEETNAME;
			}
		}
		
		return NONEXCEL;
	}
	
	/**
	 * Get a 2d array of objects in a single sheet
	 * 
	 * @param filename		filename name of excel file
	 * @param sheetIndex	sheet index
	 * @return				a 2D array of objects
	 */
	public static ArrayList<ArrayList<Object>> getSheetObject2DArray(String filename, int sheetIndex){

		Object sheet = getSheet(filename, sheetIndex);
		
		// invalid excel file
		if(sheet == null){
			System.err.println(filename + " is not excel.");
			return null;
		}
		
		// invalid excel query
		if(sheet.equals(INVALIDSHEETNAME)){
			System.err.println("Invalid sheet index:" + sheetIndex);
			return null;
		} else if(sheet.equals(NONEXCEL)){
			System.err.println(filename + " is not excel.");
			return null;
		}
		
		if(sheet instanceof HSSFSheet){
			return ReadSheet.getSheetObject2DArray((HSSFSheet)sheet);
		} else if(sheet instanceof XSSFSheet){
			return ReadSheet.getSheetObject2DArray((XSSFSheet)sheet);
		}
		
		return null;
	}
	
	/**
	 * Get a 2d array of objects in a single sheet
	 * 
	 * @param filename		filename name of excel file
	 * @param sheetIndex	sheet index
	 * @return				a 2D array of objects
	 */
	public static ArrayList<ArrayList<Object>> getSheetObject2DArray(String filename, String sheetName){

		Object sheet = getSheet(filename, sheetName);
		
		// invalid excel file
		if(sheet == null){
			System.err.println(filename + " is not excel.");
			return null;
		}
		
		// invalid excel query
		if(sheet.equals(INVALIDSHEETNAME)){
			System.err.println("Invalid sheet name:" + sheetName);
			return null;
		} else if(sheet.equals(NONEXCEL)){
			System.err.println(filename + " is not excel.");
			return null;
		}
		if(sheet instanceof HSSFSheet){
			return ReadSheet.getSheetObject2DArray((HSSFSheet)sheet);
		} else if(sheet instanceof XSSFSheet){
			return ReadSheet.getSheetObject2DArray((XSSFSheet)sheet);
		}
		
		return null;
	}
	
	public static HashMap<String, ArrayList<ArrayList<Object>>> getSheets(String filename){
		// all sheets in one HashMap<sheetName, sheet2Darray>
		HashMap<String, ArrayList<ArrayList<Object>>> sheets = new HashMap<String, ArrayList<ArrayList<Object>>>();
		ArrayList<String> names = getNames(filename);
		
		// check if the excel is valid first
		if((names == null) || (names.size() <= 0)){
			return null;
		}
		
		// loop through the excel file and store the sheets
		int i = 0;
		for(String name: names){
			ArrayList<ArrayList<Object>> sheet = getSheetObject2DArray(filename, i++);
			sheets.put(name, sheet);
		}
		
		// invalid case
		if(i == 0){
			return null;
		}
		
		return sheets;
	}
}
