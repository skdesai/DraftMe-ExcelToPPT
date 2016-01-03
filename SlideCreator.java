import java.awt.image.BufferedImage;
import java.util.Scanner;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.MalformedURLException;
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
	
	String inputFile, outputFile;
	HSLFSlideShow ppt;
	
	public static void main(String [] args) throws IOException {
		SlideCreator sc = new SlideCreator();
		sc.parseUserInput();
		sc.createTitleSlide();
		sc.parseCandidateXLSX();
        sc.writeOutput();
	    return;
	}
	
	public SlideCreator() {
		inputFile = "";
		outputFile = "";
		ppt = new HSLFSlideShow();
	}
	private void parseUserInput() {
		Scanner userInput = new Scanner(System.in);
		System.out.println("Enter the full pathname of the .xlsx workbook containing candidate data: ");
		this.inputFile = userInput.next();
		System.out.println("Enter a name for your .ppt output file: ");
		this.outputFile = userInput.next();
	}
	private void createTitleSlide() {
		HSLFSlide titleSlide = ppt.createSlide();		
		HSLFTextBox title = new HSLFTextBox();
	    HSLFTextParagraph tp1 = title.getTextParagraphs().get(0);
	    tp1.getTextRuns().get(0).setFontSize(84.);
	    tp1.setLineSpacing(120.);
	    tp1.getTextRuns().get(0).setFontFamily("Bookman Antiqua");
	    title.setText("DRAFT ME");
	    title.setAnchor(new java.awt.Rectangle(100,200,500,300));
	    title.setHorizontalCentered(true);
	    titleSlide.addShape(title);
	}
	private void parseCandidateXLSX() {	
		try {
			Workbook wb = new XSSFWorkbook(new FileInputStream(this.inputFile));
			String [] fields = null;
	        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
	            Sheet sheet = wb.getSheetAt(i);
	            for (Row row : sheet) {      	
	                int numFields = row.getPhysicalNumberOfCells();
	                if (row.getRowNum() == 0) {
	                	// Header fields
	                	fields = new String[numFields];
	                	for (Cell cell : row) {
	                		fields[cell.getColumnIndex()] = cell.toString();
	                	}
	                } else {
	                	// Row specific contents
	                	HSLFSlide slide = ppt.createSlide();
	                	HSLFTextBox shape = createDescriptionBox();

	     			    String name = "";
	                	for (Cell cell : row) {
	                		
	                		if (cell.getColumnIndex() == 0) {
	                    		HSLFTextBox title = slide.addTitle();
	                    		title.setText(cell.toString().split(" ")[0]);
	                    		name = cell.toString();
	             			    
	                		} else if (cell.getColumnIndex() == numFields-1) {
	                			HSLFPictureShape pictNew = addProfilePic(cell.toString(), name);
	                			slide.addShape(pictNew);
	                		} else {
	                			if (cell.getCellType() == 0) {
	                				shape.appendText(fields[cell.getColumnIndex()] + ": " + (int)cell.getNumericCellValue() + "\r", false);
	                			} else {
	                				shape.appendText(fields[cell.getColumnIndex()] + ": " + cell.getStringCellValue() + "\r", false);
	                			}
	                		}
	                	}
	                	slide.addShape(shape);
	                }
	            }
	        }
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return;
	}
	private HSLFPictureShape addProfilePic(String imageURL, String name) {
		HSLFPictureShape pictNew = null;
		try {
			URL profileURL = new URL(imageURL);
			BufferedImage profile = ImageIO.read(profileURL);
		    File profilePhoto = new File(name);
		    
		    ImageIO.write(profile, "jpg", profilePhoto);
		    
		    
		    // add a new picture to this slideshow and insert it in a new slide
		    HSLFPictureData pd = ppt.addPicture(profilePhoto, PictureData.PictureType.JPEG);

		    pictNew = new HSLFPictureShape(pd);
		    double WidthRescaleFactor = 250.0/((double)profile.getWidth());
		    double HeightRescaleFactor = 250.0/((double)profile.getHeight());
		    double rescaleFactor = Math.min(WidthRescaleFactor, HeightRescaleFactor);
		    int newHeight =(int) (profile.getHeight()*rescaleFactor);
		    int newWidth = (int) (profile.getWidth()*rescaleFactor);
		    pictNew.setAnchor(new java.awt.Rectangle(50, 140, newHeight, newWidth));
		   
		    return pictNew;
		} catch (IOException e) {
			e.printStackTrace();
		}
		return pictNew;
	    
	}
	private HSLFTextBox createDescriptionBox() {
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
    	shape.setAnchor(new java.awt.Rectangle(350, 140, 250, 400)); 
    	return shape;    
	}
	private void writeOutput() {
		try {
			FileOutputStream out = new FileOutputStream(outputFile);
			ppt.write(out);
			out.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

}
