package demo.poi.pp;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

public class SearchPowerPointDemo {
	/**
	 * @param args
	 */
	public static void main(String[] args) {

		try {
			XMLSlideShow ppt;
			ppt = new XMLSlideShow(new FileInputStream("slideshow.ppt"));
			
			 //append a new slide to the end
		    XSLFSlide[] slides = ppt.getSlides();
		    int i = 1;
		    for(XSLFSlide slide: slides){
		    	String title = slide.getTitle();
		    	
		    	
		    	if(title != null){
		    		System.out.println("Slide " + i + ":" + title);
		    	}
		    	else {
		    		System.out.println("Slide " + i + " has no title.");
				}
		    	
		    	XSLFTextShape[] textShapes = slide.getPlaceholders();
		    	for(XSLFTextShape textShape: textShapes){
		    		System.out.println(textShape.getText());
		    		
		    	}
		    	System.out.println();
		    	i++;
		    }
		    System.out.println(slides.length);
		    
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
