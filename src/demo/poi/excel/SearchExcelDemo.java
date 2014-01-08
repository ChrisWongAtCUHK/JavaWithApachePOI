package demo.poi.excel;

import static java.lang.System.out;
import java.util.ArrayList;
import poi.excel.SearchExcel;

/**
 * <p>
 * SearchExcelDemo
 * </p>
 * @author Chris Wong
 *
 */
public class SearchExcelDemo {
	/**
	 * Main program for demonstration
	 * @param args
	 */
	public static void main(String[] args){
		String filename = "resource\\test.xls";
		String pattern = "Chris";
	    
		ArrayList<String> result = SearchExcel.searchAllSheets(filename, pattern);
		
		for(String str: result){
			out.println(str);
		}
	}
}
