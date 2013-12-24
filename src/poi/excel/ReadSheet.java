package poi.excel;

import java.util.ArrayList;
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
	public static ArrayList<ArrayList<Object>> getSheetObject2DArray(Iterable<Row> sheet){
		ArrayList<ArrayList<Object>> object2DArray = new ArrayList<ArrayList<Object>>();
		Iterator<Row> rowIterator = sheet.iterator();
		while(rowIterator.hasNext()) {
			ArrayList<Object> objects = new ArrayList<Object>();
		    Row row = rowIterator.next();
		     
		    //For each row, iterate through each columns
		    Iterator<Cell> cellIterator = row.cellIterator();
		    while(cellIterator.hasNext()) {
		         
		        Cell cell = cellIterator.next();
		         
		        switch(cell.getCellType()) {
		            case Cell.CELL_TYPE_BOOLEAN:
		                objects.add(cell.getBooleanCellValue());
		                break;
		            case Cell.CELL_TYPE_NUMERIC:
		                objects.add(cell.getNumericCellValue());
		                break;
		            case Cell.CELL_TYPE_STRING:
		                objects.add(cell.getStringCellValue());
		                break;
		        }
		    }
		    object2DArray.add(objects);
		}
		
		return object2DArray;
	}
}
