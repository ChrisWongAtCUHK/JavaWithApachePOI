package poi.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
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
 * List the names of sheet in a single xls/xlsx
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
	 * listSheets
	 */
	public static void listSheets(){
		
	}
}
