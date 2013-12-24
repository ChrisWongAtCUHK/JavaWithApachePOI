package poi.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

//For .xls, libraries included: poi-3.9-20121203.jar
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

//For .xlsx, so many additional libraries must included: ooxml-lib/dom4j-1.6.1.jar, ooxml-lib/xmlbeans-2.3.0.jar, poi-ooxml-3.9-20121203.jar, poi-ooxml-schemas-3.9-20121203.jar
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
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
	final static String XLSFILENAME = "resource\\sheets.xls";
	final static String XLSXFILENAME = "resource\\sheets.xlsx";
	
	public static void main(String[] args){
		try {
		     
		    FileInputStream file = new FileInputStream(new File(XLSFILENAME));
		    FileInputStream xfile = new FileInputStream(new File(XLSXFILENAME));
		     
		    //Get the workbook instance for XLS file 
		    HSSFWorkbook workbook = new HSSFWorkbook(file);
		 
		    //Get the workbook instance for XLSX file 
		    XSSFWorkbook xworkbook = new XSSFWorkbook(xfile);
		    
		    //Get first sheet from the workbook, for XLS file
		    HSSFSheet sheet = workbook.getSheetAt(0);
		    
		    //Get first sheet from the workbook, for XLSX file
		    XSSFSheet xsheet = xworkbook.getSheetAt(0);
		    
		    //Iterate through each rows from first sheet
		    out.println("--------------------XLS file sheets------------------------");
		    ReadSheet.sheetIterate(sheet);
		    out.println("--------------------XLSX file sheets------------------------");
		    ReadSheet.sheetIterate(xsheet);
		    file.close();
		     
		} catch (FileNotFoundException e) {
		    e.printStackTrace();
		} catch (IOException e) {
		    e.printStackTrace();
		}
	}
	

	/**
	 * Get names of sheets
	 * 
	 * @return names of sheets
	 */
	public static ArrayList<String> getNames(String filename){
		// Check the type of excel file
		// Use HSSF or XSSF accordingly
		String excelType = ExcelFile.excelType(filename);
		
		// ArrayList to store names of sheets
		ArrayList<String> namesArrayList = new ArrayList<String>();
		
		if(excelType == null){
			// neither xls nor xlsx
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
			    	namesArrayList.add(workbook.getSheetName(i));
			    }
			    
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
			    int numberOfSheets = xworkbook.getNumberOfSheets();
			    
			    // loop through the workbook, add the name of sheet to the arraylist one by one
			    for(int i = 0; i < numberOfSheets; i++){
			    	
			    	namesArrayList.add(xworkbook.getSheetName(i));
			    }
			    
			} catch (FileNotFoundException e) {
			    e.printStackTrace();
			} catch (IOException e) {
			    e.printStackTrace();
			}
		}
		
		return namesArrayList;
	}
}
