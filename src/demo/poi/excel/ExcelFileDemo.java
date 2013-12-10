package demo.poi.excel;

// Self-defined class
import poi.excel.ExcelFile;
import static java.lang.System.out;

/**
 * <p>
 *  ExcelFileDemo
 * </p>
 * @author Chris Wong
 * A demo program for poi.excel.ExcelFile
 */
public class ExcelFileDemo {

	static String txtFileName = "resource\\test.txt";
	static String xlsFileName = "resource\\test.xls";
	static String xlsxFileName = "resource\\test.xlsx";
	
	public static void main(String[] args){
		out.println(ExcelFile.excelType(txtFileName));
		out.println(ExcelFile.excelType(xlsFileName));
		out.println(ExcelFile.excelType(xlsxFileName));
	}
}
