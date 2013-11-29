package poi.excel;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.Date;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

// For .xls, libraries included: poi-3.9-20121203.jar
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

// For .xlsx, so many additional libraries must included: ooxml-lib/dom4j-1.6.1.jar, ooxml-lib/xmlbeans-2.3.0.jar, poi-ooxml-3.9-20121203.jar, poi-ooxml-schemas-3.9-20121203.jar
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class WriteDemo {
	// http://viralpatel.net/blogs/java-read-write-excel-file-apache-poi/
	public static void main(String[] args){
		
		// Create the spreadsheet
		
		// For .xls
		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheet = workbook.createSheet("Sample sheet");
		
		// For .xlsx
		XSSFWorkbook xworkbook = new XSSFWorkbook();
		XSSFSheet xsheet = xworkbook.createSheet("Sample xsheet");
		
		// Insert the data to TreeMap, which have already sorted the data
		Map<String, Object[]> data = new TreeMap<String, Object[]>();
		data.put("1", new Object[] {"Emp No.", "Name", "Salary"});
		data.put("2", new Object[] {1d, "John", 1500000d});
		data.put("3", new Object[] {2d, "Sam", 800000d});
		data.put("4", new Object[] {3d, "Dean", 700000d});
		 
		// For XLS file
		sheetWrite(data, sheet);
		
		// For XLSX file
		sheetWrite(data, xsheet);
		
		try {
			
			// For .xls
			FileOutputStream out = new FileOutputStream(new File("resource\\test.xls"));
			workbook.write(out);
			out.close();
			System.out.println("XLS written successfully..");
		    
			// For .xlsx
		    FileOutputStream xout = new FileOutputStream(new File("resource\\test.xlsx"));
		    xworkbook.write(xout);
		    xout.close();
		    System.out.println("XLSX written successfully..");
		    
		     
		} catch (FileNotFoundException e) {
		    e.printStackTrace();
		} catch (IOException e) {
		    e.printStackTrace();
		}
	}
	
	// Apply the data to the spreadsheet
	public static void sheetWrite(Map<String, Object[]> data, Iterable<Row> iterable){
		
		Set<String> keyset = data.keySet();
		int rownum = 0;
		
		for (String key : keyset) {
			
			// Each row
		    Row row = null;
		    
		    // Reflect: http://godleon.blogspot.hk/2007/09/class-class-java-class-class-jvm-class.html
			try {
				// Get the class, IMPORTANT: getName()
				Class<?> c = Class.forName(iterable.getClass().getName());
				
				// Cast the interface to object
				Object obj = c.cast(iterable);
				
				// Specify the parameter(s)  for method
				Class<?>[] clazz = {Integer.TYPE};
				
				// Create method with method name and parameters
				Method createRow = c.getMethod("createRow", clazz);
				
				// Specify the value(s) of parameter(s)
				Object[] args = {new Integer(rownum++)};
				
				// Invoke the method
				row = (Row)createRow.invoke(obj, args);
			} catch (ClassNotFoundException e) {
				// For Class.forName
				e.printStackTrace();
			} catch (SecurityException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (NoSuchMethodException e) {
				// // For getMethod
				e.printStackTrace();
			} catch (IllegalArgumentException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (IllegalAccessException e) {
				// For invoke method
				e.printStackTrace();
			} catch (InvocationTargetException e) {
				// For invoke method
				e.printStackTrace();
			}
		    
		    Object [] objArr = data.get(key);
		    int cellnum = 0;
		    for (Object obj : objArr) {
		    	
		    	// Columns
		        Cell cell = row.createCell(cellnum++);
		        if(obj instanceof Date) 
		            cell.setCellValue((Date)obj);
		        else if(obj instanceof Boolean)
		            cell.setCellValue((Boolean)obj);
		        else if(obj instanceof String)
		            cell.setCellValue((String)obj);
		        else if(obj instanceof Double)
		            cell.setCellValue((Double)obj);
		    }
		}
	}
}
