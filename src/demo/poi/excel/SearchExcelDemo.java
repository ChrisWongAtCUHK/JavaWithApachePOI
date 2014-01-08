package demo.poi.excel;

import static java.lang.System.out;

import java.util.ArrayList;

// 
import file.io.FileList;

import poi.excel.SearchExcel;


/**
 * @author Chris Wong
 *
 */
public class SearchExcelDemo {
	public static void main(String[] args){
		String filename = "resource\\test.xls";
		String pattern = "Chris";
	    
		ArrayList<String> result = SearchExcel.searchAllSheets(filename, pattern);
		
		for(String str: result){
			out.println(str);
		}
		String path = "D:\\tmp";
		if(args.length > 0){
			path = args[0];
		}
		ArrayList<String> files = FileList.getFiles(path);
		for(String file: files){
			out.format("%s%n", file);
		}
	    
	}
}
