package demo.poi.pp;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;

/**
 * @author Chris Wong
 * <p>
 * <a href="http://poi.apache.org/slideshow/xslf-cookbook.html">ReadPresentation</a>
 * </p>
 */
public class ReadPresentation {
	public static void main(String[] args){
		
		try {
			XMLSlideShow ppt;
			ppt = new XMLSlideShow(new FileInputStream("slideshow.ppt"));
			
			 //append a new slide to the end
		    XSLFSlide blankSlide = ppt.createSlide();
		    
		    // write to a file
			FileOutputStream out;
			try {
				out = new FileOutputStream("slideshow.ppt");
				ppt.write(out);
				out.close();
			} catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	   
	}
}
