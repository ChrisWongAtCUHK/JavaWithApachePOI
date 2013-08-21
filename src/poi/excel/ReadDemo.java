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

public class ReadDemo {
	
	public static void main(String[] args){
		try {
		     
		    FileInputStream file = new FileInputStream(new File("test.xls"));
		    FileInputStream xfile = new FileInputStream(new File("test.xlsx"));
		     
		    //Get the workbook instance for XLS file 
		    HSSFWorkbook workbook = new HSSFWorkbook(file);
		 
		    //Get the workbook instance for XLSX file 
		    XSSFWorkbook xworkbook = new XSSFWorkbook(xfile);
		    
		    //Get first sheet from the workbook, for XLS file
		    HSSFSheet sheet = workbook.getSheetAt(0);
		    
		    //Get first sheet from the workbook, for XLSX file
		    XSSFSheet xsheet = xworkbook.getSheetAt(0);
		    
		    //Iterate through each rows from first sheet
		    System.out.println("--------------------XLS file------------------------");
		    sheetIterate(sheet);
		    System.out.println("--------------------XLSX file------------------------");
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
