package poi.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.security.GeneralSecurityException;
import java.util.Iterator;

//For .xls, libraries included: poi-3.9-20121203.jar
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

//For .xlsx, so many additional libraries must included: ooxml-lib/dom4j-1.6.1.jar, ooxml-lib/xmlbeans-2.3.0.jar, poi-ooxml-3.9-20121203.jar, poi-ooxml-schemas-3.9-20121203.jar
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//For protected xls/xlsx
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class ReadProtectedDemo {
	final static String PASSWORD = "123456";
	public static void main(String[] args){
		try {
			
			InputStream file = new FileInputStream(new File("resource\\protected.xls"));
		    FileInputStream xfile = new FileInputStream(new File("resource\\protected.xlsx"));
		    org.apache.poi.hssf.record.crypto.Biff8EncryptionKey.setCurrentUserPassword(PASSWORD);
		    
		    POIFSFileSystem fs = new POIFSFileSystem(xfile);
		    EncryptionInfo info = new EncryptionInfo(fs);
		    Decryptor d = Decryptor.getInstance(info);
		    
		    InputStream dataStream = null;
		    try {
	
				 if (!d.verifyPassword(PASSWORD)) {
				        throw new RuntimeException("Unable to process: document is encrypted");
				    }

				    dataStream = d.getDataStream(fs);

				    // parse dataStream
			} catch (GeneralSecurityException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		    
		    //Get the workbook instance for XLS file 
		    HSSFWorkbook workbook = new HSSFWorkbook(file);
		 
		    //Get the workbook instance for XLSX file 
		    XSSFWorkbook xworkbook = new XSSFWorkbook(dataStream);
		    
		    //Get first sheet from the workbook, for XLS file
		    HSSFSheet sheet = workbook.getSheetAt(0);
		    
		    //Get first sheet from the workbook, for XLSX file
		    XSSFSheet xsheet = xworkbook.getSheetAt(0);
		    
		    //Iterate through each rows from first sheet
		    System.out.println("--------------------Protected XLS file------------------------");
		    sheetIterate(sheet);
		    System.out.println("--------------------Protected XLSX file------------------------");
		    sheetIterate(xsheet);
		    file.close();
		     
		} catch (FileNotFoundException e) {
		    e.printStackTrace();
		} catch (IOException e) {
		    e.printStackTrace();
		}
	}
	
	// Iterate the rows of the sheet from XLS/XLS file
	public static void sheetIterate(Iterable<Row> sheet){
		Iterator<Row> rowIterator = sheet.iterator();
		while(rowIterator.hasNext()) {
		    Row row = rowIterator.next();
		     
		    //For each row, iterate through each columns
		    Iterator<Cell> cellIterator = row.cellIterator();
		    while(cellIterator.hasNext()) {
		         
		        Cell cell = cellIterator.next();
		        
		        switch(cell.getCellType()) {
		            case Cell.CELL_TYPE_BOOLEAN:
		                System.out.print(cell.getBooleanCellValue() + "\t\t");
		                break;
		            case Cell.CELL_TYPE_NUMERIC:
		                System.out.print(cell.getNumericCellValue() + "\t\t");
		                break;
		            case Cell.CELL_TYPE_STRING:
		                System.out.print(cell.getStringCellValue() + "\t\t");
		                break;
		        }
		        
		    }
		    System.out.println();
		}
	}
}
