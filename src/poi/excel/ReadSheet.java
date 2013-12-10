package poi.excel;

import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

/**
 * <p>
 *  ReadSheet
 * </p>
 *  Read the information or content of a single sheet
 * @author Chris Wong
 *
 */
public class ReadSheet {
	
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
	
	// Iterate the rows of the sheet from XLS/XLS file
	public static Object[][] getSheet(Iterable<Row> sheet){
		Object[][] objects = new Object[99][99];
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
		
		return objects;
	}
}
