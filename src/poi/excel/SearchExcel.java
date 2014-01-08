package poi.excel;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Map.Entry;

/**
 * <p>
 *  SearchExcel searhes content in excel file
 * </p>
 * @author Chris Wong
 *
 */
public class SearchExcel {
	public static ArrayList<String> searchAllSheets(String filename, String pattern){
		// the ArrayList with filename and sheetname
		ArrayList<String> filenamesAndSheetnames = new ArrayList<String>();
		
		// get all contents 
		HashMap<String, ArrayList<ArrayList<Object>>> sheets = ListSheets.getAllSheets(filename);
		
		// invalid excel file
		if(sheets == null){
			return filenamesAndSheetnames;
		}
		
		Iterator<Entry<String, ArrayList<ArrayList<Object>>>> it = sheets.entrySet().iterator();
	    while (it.hasNext()) {
	        Map.Entry<String, ArrayList<ArrayList<Object>>> pairs = (Map.Entry<String, ArrayList<ArrayList<Object>>>)it.next();
	        
	        // key is the sheet name
	        ArrayList<ArrayList<Object>> sheetObject2DArray = pairs.getValue();
	        
	        // avoid invalid cases
	        if(sheetObject2DArray != null){
	        	// value is sheet content
		        for(ArrayList<Object> objects: sheetObject2DArray){
					for(Object object: objects){
						// pattern matching
						if(object.toString().matches(".*"+ pattern +".*")){
							filenamesAndSheetnames.add(filename + "," + pairs.getKey());
						}
					}
					
				}
	        }

	        it.remove(); // avoids a ConcurrentModificationException
	    }
	    
	    return filenamesAndSheetnames;
	}
}
