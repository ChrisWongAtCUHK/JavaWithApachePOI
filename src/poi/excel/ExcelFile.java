package poi.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

//For .xls, libraries included: poi-3.9-20121203.jar
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//For .xlsx, so many additional libraries must included: ooxml-lib/dom4j-1.6.1.jar, ooxml-lib/xmlbeans-2.3.0.jar, poi-ooxml-3.9-20121203.jar, poi-ooxml-schemas-3.9-20121203.jar
import org.apache.poi.POIXMLException;

/**
 * <p>
 * 	ExcelFile
 * </p> 
 * Check if a file is valid for using HSSF or XSSF
 * @author Chris Wong
 *
 */
public class ExcelFile {

	final static String XLSTYPE = "xls";
	final static String XLSXTYPE = "xlsx";
	
	// Check if a file is excel file or not, null is invalid
	public static String excelType(String fileName){


		// check the file by name, case by case
		if(isXls(fileName) == true){
			// xls file
			return XLSTYPE;
		} else if(isXlsx(fileName) == true){
			// xlsx file
			return XLSXTYPE;
		}
		
		return null;
	}
	
	// to check if it is xls file
	public static boolean isXls(String fileName){
		try {
			File file = new File(fileName);
			FileInputStream fis = new FileInputStream(file);
		    
		    //Get the workbook instance for XLS file 
		    new HSSFWorkbook(fis);

		    return true;
		} catch (FileNotFoundException e) {
		    e.printStackTrace();
		} catch (IOException e) {
			return false;
		} catch (OfficeXmlFileException oxfe){
			// may be xlsx file
			return false;
		} catch (POIXMLException pe){
			// other type of file
			return false;
		}
		
		return false;
	}
	
	// to check if it is xlsx 
	public static boolean isXlsx(String fileName){
		try {
			File file = new File(fileName);
			FileInputStream fis = new FileInputStream(file);
			
		    //Get the workbook instance for XLS file 
		    new XSSFWorkbook(fis);
		    
		    return true;
		} catch (FileNotFoundException e) {
		    e.printStackTrace();
		} catch (IOException e) {
			return false;
		} catch (OfficeXmlFileException oxfe){
			// may be xls file
			return false;
		} catch (POIXMLException pe){
			// other type of file
			return false;
		}
		
		return false;
	}
	
}
