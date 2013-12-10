package demo.poi.excel;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

// For .xls, libraries included: poi-3.9-20121203.jar
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

// For .xlsx, so many additional libraries must included: ooxml-lib/dom4j-1.6.1.jar, ooxml-lib/xmlbeans-2.3.0.jar, poi-ooxml-3.9-20121203.jar, poi-ooxml-schemas-3.9-20121203.jar
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

// Self-defined class
import poi.excel.ReadSheet;

/**
 * <p>
 *  ReadDemo
 * </p>
 * @author Chris Wong
 * A demo program for poi.excel.ReadSheet
 */
public class ReadDemo {
	
	static String xlsFileName = "resource\\test.xls";
	static String xlsxFileName = "resource\\test.xlsx";
	
	public static void main(String[] args){
		try {
		    
			File file = new File(xlsFileName);
		    FileInputStream fis = new FileInputStream(new File(xlsFileName));
		    FileInputStream xfis = new FileInputStream(new File(xlsxFileName));
		     
		    //Get the workbook instance for XLS file 
		    HSSFWorkbook workbook = new HSSFWorkbook(fis);
		 
		    //Get the workbook instance for XLSX file 
		    XSSFWorkbook xworkbook = new XSSFWorkbook(xfis);
		    
		    //Get first sheet from the workbook, for XLS file
		    HSSFSheet sheet = workbook.getSheetAt(0);
		    
		    //Get first sheet from the workbook, for XLSX file
		    XSSFSheet xsheet = xworkbook.getSheetAt(0);
		    
		    //Iterate through each rows from first sheet
		    System.out.println("--------------------XLS file------------------------");
		    ReadSheet.sheetIterate(sheet);
		    System.out.println("--------------------XLSX file------------------------");
		    ReadSheet.sheetIterate(xsheet);
		    fis.close();
		     
		} catch (FileNotFoundException e) {
		    e.printStackTrace();
		} catch (IOException e) {
		    e.printStackTrace();
		}
	}
	
}
