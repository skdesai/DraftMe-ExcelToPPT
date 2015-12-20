import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;

import javax.imageio.ImageIO;

import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.hslf.usermodel.HSLFTextBox;
import org.apache.poi.hslf.usermodel.HSLFTextParagraph;
import org.apache.poi.hslf.usermodel.HSLFTextRun;
import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hslf.usermodel.HSLFPictureData;
import org.apache.poi.hslf.usermodel.HSLFPictureShape;
import org.apache.poi.hslf.usermodel.HSLFSlide;

public class SlideCreator {
	//create a new empty slide show
	
	public static void main(String [] args) throws IOException {
		HSLFSlideShow ppt = new HSLFSlideShow();
		
		String [] fields = null;
		Workbook wb = new XSSFWorkbook(new FileInputStream("retreat.xlsx"));
		
		HSLFSlide titleSlide = ppt.createSlide();		
		HSLFTextBox shape1 = new HSLFTextBox();
	    HSLFTextParagraph tp1 = shape1.getTextParagraphs().get(0);
	    tp1.getTextRuns().get(0).setFontSize(84.);
	    tp1.setLineSpacing(120.);
	    tp1.getTextRuns().get(0).setFontFamily("Bookman Antiqua");
	    shape1.setText("DRAFT ME");
	    shape1.setAnchor(new java.awt.Rectangle(100,200,500,300));
	    shape1.setHorizontalCentered(true);
	    titleSlide.addShape(shape1);
	   
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            Sheet sheet = wb.getSheetAt(i);
            System.out.println(wb.getSheetName(i));
            for (Row row : sheet) {      	
                System.out.println("rownum: " + row.getRowNum());
                int numFields = row.getPhysicalNumberOfCells();
                if (row.getRowNum() == 0) {
                	// Headers
                	fields = new String[numFields];
                	for (Cell cell : row) {
                		fields[cell.getColumnIndex()] = cell.toString();
                	}
                } else {
                	// Row specific contents
                	HSLFSlide slide = ppt.createSlide();
           			HSLFTextBox shape = new HSLFTextBox();
     			    HSLFTextParagraph tp = shape.getTextParagraphs().get(0);
     			    tp.getTextRuns().get(0).setFontSize(20.);
     			    tp.setLineSpacing(120.);
     			    
     			    tp.setBullet(true);
     			    tp.setBulletChar('\u2022'); //bullet character
     			    tp.setIndent(0.);  //bullet offset
     			    tp.setLeftMargin(20.);   //text offset (should be greater than bullet offset)
     			    HSLFTextRun rt = tp.getTextRuns().get(0);
     			    shape.setText("");
     			    

     			    String name = "";
     			    //shape.getText().length();
                	for (Cell cell : row) {
                		
                		if (cell.getColumnIndex() == 0) {
                    		HSLFTextBox title = slide.addTitle();
                    		title.setText(cell.toString().split(" ")[0]);
                    		name = cell.toString();
             			    
                		} 
                		File profilePhoto = new File("logo.jpg");
                		HSLFPictureData pd = ppt.addPicture(profilePhoto, PictureData.PictureType.JPEG);
                		HSLFPictureShape pictNew = new HSLFPictureShape(pd);
                		//pictNew.setAnchor(new java.awt.Rectangle(50, 140));
                		pictNew.setAnchor(new java.awt.Rectangle(25,440,250,100));
                		slide.addShape(pictNew);
//                		else if (cell.getColumnIndex() == numFields-1) {
//                			URL profileURL = new URL(cell.toString());
//             			    BufferedImage profile = ImageIO.read(profileURL);
//             			    File profilePhoto = new File(name);
//             			    
//             			    ImageIO.write(profile, "jpg", profilePhoto);
//             			    
//             			    
//             			    // add a new picture to this slideshow and insert it in a new slide
//             			    HSLFPictureData pd = ppt.addPicture(profilePhoto, PictureData.PictureType.JPEG);
//
//             			    HSLFPictureShape pictNew = new HSLFPictureShape(pd);
//             			    double WidthRescaleFactor = 250.0/((double)profile.getWidth());
//             			    double HeightRescaleFactor = 250.0/((double)profile.getHeight());
//             			    double rescaleFactor = Math.min(WidthRescaleFactor, HeightRescaleFactor);
//             			    int newHeight =(int) (profile.getHeight()*rescaleFactor);
//             			    int newWidth = (int) (profile.getWidth()*rescaleFactor);
//             			    pictNew.setAnchor(new java.awt.Rectangle(50, 140, newHeight, newWidth));
//             			   
//             			    slide.addShape(pictNew);
//                		} else {
//                			if (cell.getCellType() == 0) {
//                				shape.appendText(fields[cell.getColumnIndex()] + ": " + (int)cell.getNumericCellValue() + "\r", false);
//                			} else {
//                				shape.appendText(fields[cell.getColumnIndex()] + ": " + cell.getStringCellValue() + "\r", false);
//                			}
//                			
//                		}
                	}
                	shape.setAnchor(new java.awt.Rectangle(350, 140, 250, 400)); 
                	slide.addShape(shape);
                }
            }
        }

	    
	    //save changes in a file
	    FileOutputStream out;
		out = new FileOutputStream("retreat.ppt");
		ppt.write(out);
		out.close();

	
	    return;
	}

}
