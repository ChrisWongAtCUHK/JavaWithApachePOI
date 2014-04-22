package demo.poi.pp;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xslf.usermodel.SlideLayout;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

/**
 * @author Chris Wong
 * <p>
 * <a href="http://poi.apache.org/slideshow/xslf-cookbook.html">CreateNewPresentation</a>
 * </p>
 */
public class CreateNewPresentation {
	public static void main(String[] args) {
		XMLSlideShow ppt = new XMLSlideShow();
		
		// create a blank slide
		XSLFSlide slide = ppt.createSlide();
		
		// there can be multiple masters each referencing a number of layouts
	    // for demonstration purposes we use the first (default) slide master
	    XSLFSlideMaster defaultMaster = ppt.getSlideMasters()[0];

	    // title slide
	    XSLFSlideLayout titleLayout = defaultMaster.getLayout(SlideLayout.TITLE);
	    // fill the placeholders
	    XSLFSlide slide1 = ppt.createSlide(titleLayout);
	    XSLFTextShape title1 = slide1.getPlaceholder(0);
	    title1.setText("First Title");

	    // title and content
	    XSLFSlideLayout titleBodyLayout = defaultMaster.getLayout(SlideLayout.TITLE_AND_CONTENT);
	    XSLFSlide slide2 = ppt.createSlide(titleBodyLayout);

	    XSLFTextShape title2 = slide2.getPlaceholder(0);
	    title2.setText("Second Title");

	    XSLFTextShape body2 = slide2.getPlaceholder(1);
	    body2.clearText(); // unset any existing text
	    body2.addNewTextParagraph().addNewTextRun().setText("First paragraph");
	    body2.addNewTextParagraph().addNewTextRun().setText("Second paragraph");
	    body2.addNewTextParagraph().addNewTextRun().setText("Third paragraph");
	    
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
		
	}
}
