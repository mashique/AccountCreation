


import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.awt.AWTException;
import java.awt.FlowLayout;
import java.awt.Rectangle;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.event.InputEvent;
import java.awt.event.KeyEvent;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.StringSelection;
import java.awt.*;
import java.awt.event.*;
import java.awt.image.BufferedImage;

import javax.imageio.ImageIO;
import javax.swing.JFrame;
import javax.swing.JLabel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



	public class L_TWC_QAHA_HI_TriplePlay {
    
	    public static final int SleepBtwClicks = 500;
	    public static final int NewPageLoad = 4000;
	    public static int MAX_Y = 440;
	    public static int MAX_X = 110;
	    public static int StartExecution = 2000;
	    public static String FirstName;
	    public static String LocationID;
	    public static String VideoSN;
	    public static String VideoSN2;
	    public static String HSDSN;
	    public static String AccountNumber;
	    public static String OrderID;
	    static boolean ImageMatched;
	    
	    public static void main(String... args) throws Exception {
	    Robot robot = new Robot();

		FileInputStream f=new FileInputStream("C:\\Driver\\AutomatedAccountCreation.xlsx");
		System.out.println("value in f is "+f);
		XSSFWorkbook wb = new XSSFWorkbook(f);
		System.out.println("value in wb is "+wb);
		XSSFSheet s = wb.getSheet("Sheet1");
		System.out.println("value in s is "+s);
		int totalRow=s.getPhysicalNumberOfRows();
		
		System.out.println("number of rows in spreadsheet "+totalRow);
		//Making change in for loop to understand values saved in object
		for(int row=1; row<totalRow; row++){
			try{
			Cell cell1=s.getRow(row).getCell(0);
			LocationID=cell1.getStringCellValue();
			Cell cell2=s.getRow(row).getCell(1);
			FirstName=cell2.getStringCellValue();
			Cell cell3=s.getRow(row).getCell(2);
			HSDSN=cell3.getStringCellValue();
			Cell cell4=s.getRow(row).getCell(3);
			VideoSN=cell4.getStringCellValue();
			Cell cell5=s.getRow(row).getCell(4);
			VideoSN2=cell5.getStringCellValue();
			
		Thread.sleep(StartExecution);
//		Toolkit toolkit = Toolkit.getDefaultToolkit();
//		Clipboard clipboard = toolkit.getSystemClipboard();

		StringSelection selection = new StringSelection(LocationID);
		Clipboard clipboard1 = Toolkit.getDefaultToolkit().getSystemClipboard();
		clipboard1.setContents(selection, selection);
	    Thread.sleep(SleepBtwClicks);
	    
		Thread.sleep(NewPageLoad);
		//Thread.sleep(SleepBtwClicks);
		robot.keyPress(KeyEvent.VK_ALT);
		robot.keyPress(KeyEvent.VK_C);
		robot.keyRelease(KeyEvent.VK_C);
		robot.keyRelease(KeyEvent.VK_ALT);
		
		Thread.sleep(NewPageLoad);
		//Thread.sleep(SleepBtwClicks);
		robot.keyPress(KeyEvent.VK_ALT);
		robot.keyPress(KeyEvent.VK_C);
		robot.keyRelease(KeyEvent.VK_C);
		robot.keyRelease(KeyEvent.VK_ALT);
		//Thread.sleep(SleepBtwClicks);
		//Thread.sleep(SleepBtwClicks);
		
	  /*  //Move Cursor to hit Clear
		MAX_X = 1320;
		MAX_Y = 676;
	    robot.mouseMove(MAX_X, MAX_Y);
	    Thread.sleep(SleepBtwClicks);
	    robot.mousePress(InputEvent.BUTTON1_MASK);
	    robot.mouseRelease(InputEvent.BUTTON1_MASK);
		*/
//*Validate to be on correct page to proceed - First page
		//ImageMatched = false;
		ImageMatched = true;
		while (!ImageMatched){

//*Save the current image at "Search" button position
	       try {
	          /*
	           * Let the program wait for 5 seconds to allow you to open the
	           * window whose screenshot has to be captured
	           */
	    	   System.out.println("first BP");
	           Thread.sleep(NewPageLoad);
	           Robot robot2 = new Robot();
	           String fileName = "C://Driver//Search_Button_Current.jpg";
	 
	           // Define an area of size 500*400 starting at coordinates (10,50)
	           //Rectangle rectArea = new Rectangle(24, 218, 56, 12);
	           Rectangle rectArea = new Rectangle(1083, 664, 88, 27);
	           BufferedImage screenFullImage = robot2.createScreenCapture(rectArea);
	           ImageIO.write(screenFullImage, "jpg", new File(fileName));
	 

	       } catch (IOException | InterruptedException ex) {
	                System.err.println(ex);
	       }
		
//Compare the image with saved image - 
	       BufferedImage imgA = null; 
	       BufferedImage imgB = null; 

	       try
	       { 
	           File fileA = new File("C://Driver//Search_Button_Current.jpg"); 
	           File fileB = new File("C://Driver//Search_Button.jpg"); 

	           imgA = ImageIO.read(fileA); 
	           imgB = ImageIO.read(fileB); 
	       } 
	       catch (IOException e) 
	       { 
	           System.out.println(e); 
	       } 
	       int width1 = imgA.getWidth(); 
	       int width2 = imgB.getWidth(); 
	       int height1 = imgA.getHeight(); 
	       int height2 = imgB.getHeight(); 

	       System.out.println("width and hights"+width1+width2+height1+height2);
	       if ((width1 == width2) && (height1 == height2)) 
	       {
		       System.out.println("inside first if");
	    	   long difference = 0; 
	           for (int y = 0; y < height1; y++) 
	           { 
	               for (int x = 0; x < width1; x++) 
	               { 
	                   int rgbA = imgA.getRGB(x, y); 
	                   int rgbB = imgB.getRGB(x, y); 
	                   int redA = (rgbA >> 16) & 0xff; 
	                   int greenA = (rgbA >> 8) & 0xff; 
	                   int blueA = (rgbA) & 0xff; 
	                   int redB = (rgbB >> 16) & 0xff; 
	                   int greenB = (rgbB >> 8) & 0xff; 
	                   int blueB = (rgbB) & 0xff; 
	                   difference += Math.abs(redA - redB); 
	                   difference += Math.abs(greenA - greenB); 
	                   difference += Math.abs(blueA - blueB); 

	               } 
	           }
	           System.out.println("difference is"+difference);
	           //sometimes images don;t match exactly, value 1000 is high but it turns out to be a small percentage
	           if (difference <= 4000){
	        	   System.out.println("Image matched");
	        	   ImageMatched = true;
	           }
	       }
	       else
	       { 
	    	  // ImageMatched = false;
	    	   Thread.sleep(NewPageLoad);
	    	   System.out.println("first BP3");
	       }
		}   
	       
	    //Move Cursor to enter LocationID
		MAX_X = 755;
		MAX_Y = 427;
	    robot.mouseMove(MAX_X, MAX_Y);
	    Thread.sleep(SleepBtwClicks);
	    robot.mousePress(InputEvent.BUTTON1_MASK);
	    robot.mouseRelease(InputEvent.BUTTON1_MASK);
	    robot.mousePress(InputEvent.BUTTON1_MASK);
	    robot.mouseRelease(InputEvent.BUTTON1_MASK);
	    Thread.sleep(SleepBtwClicks);
		robot.keyPress(KeyEvent.VK_CONTROL);
		robot.keyPress(KeyEvent.VK_V);
		robot.keyRelease(KeyEvent.VK_V);
		robot.keyRelease(KeyEvent.VK_CONTROL);
		Thread.sleep(SleepBtwClicks);
		
	    //Hit Enter 
		Thread.sleep(SleepBtwClicks);
		robot.keyPress(KeyEvent.VK_ENTER);
		robot.keyRelease(KeyEvent.VK_ENTER);
		
	    //New Connect - Alt + O
	    Thread.sleep(NewPageLoad);
		robot.keyPress(KeyEvent.VK_ALT);
		robot.keyPress(KeyEvent.VK_O);
		robot.keyRelease(KeyEvent.VK_O);
		robot.keyRelease(KeyEvent.VK_ALT);
		
		//Enter Last Name
		Thread.sleep(NewPageLoad);
	      robot.keyPress(KeyEvent.VK_T);
	      robot.keyRelease(KeyEvent.VK_T);
	      robot.keyPress(KeyEvent.VK_A);
	      robot.keyRelease(KeyEvent.VK_A);
	      robot.keyPress(KeyEvent.VK_M);
	      robot.keyRelease(KeyEvent.VK_M);
	      robot.keyPress(KeyEvent.VK_T);
	      robot.keyRelease(KeyEvent.VK_T);
	      robot.keyPress(KeyEvent.VK_O);
	      robot.keyRelease(KeyEvent.VK_O);
	      robot.keyPress(KeyEvent.VK_O);
	      robot.keyRelease(KeyEvent.VK_O);
	      robot.keyPress(KeyEvent.VK_L);
	      robot.keyRelease(KeyEvent.VK_L);
		
			//Enter First Name
	      Thread.sleep(SleepBtwClicks);
		selection = new StringSelection(FirstName);
		clipboard1 = Toolkit.getDefaultToolkit().getSystemClipboard();
		clipboard1.setContents(selection, selection);
	    Thread.sleep(SleepBtwClicks);
	      robot.keyPress(KeyEvent.VK_TAB);
	      robot.keyRelease(KeyEvent.VK_TAB);
		    Thread.sleep(SleepBtwClicks);
			robot.keyPress(KeyEvent.VK_CONTROL);
			robot.keyPress(KeyEvent.VK_V);
			robot.keyRelease(KeyEvent.VK_V);
			robot.keyRelease(KeyEvent.VK_CONTROL);
			Thread.sleep(SleepBtwClicks);
						
			//Enter Telephone number
		      robot.keyPress(KeyEvent.VK_TAB);
		      robot.keyRelease(KeyEvent.VK_TAB);
			Thread.sleep(SleepBtwClicks);
		      robot.keyPress(KeyEvent.VK_3);
		      robot.keyRelease(KeyEvent.VK_3);
		      robot.keyPress(KeyEvent.VK_1);
		      robot.keyRelease(KeyEvent.VK_1);
		      robot.keyPress(KeyEvent.VK_4);
		      robot.keyRelease(KeyEvent.VK_4);
		      robot.keyPress(KeyEvent.VK_9);
		      robot.keyRelease(KeyEvent.VK_9);
		      robot.keyPress(KeyEvent.VK_9);
		      robot.keyRelease(KeyEvent.VK_9);
		      robot.keyPress(KeyEvent.VK_9);
		      robot.keyRelease(KeyEvent.VK_9);
		      robot.keyPress(KeyEvent.VK_9);
		      robot.keyRelease(KeyEvent.VK_9);
		      robot.keyPress(KeyEvent.VK_9);
		      robot.keyRelease(KeyEvent.VK_9);
		      robot.keyPress(KeyEvent.VK_9);
		      robot.keyRelease(KeyEvent.VK_9);
		      robot.keyPress(KeyEvent.VK_9);
		      robot.keyRelease(KeyEvent.VK_9);
		
				
	    //Hit Enter - Thrice
		Thread.sleep(SleepBtwClicks);
		robot.keyPress(KeyEvent.VK_ENTER);
		robot.keyRelease(KeyEvent.VK_ENTER);
		Thread.sleep(SleepBtwClicks);
		robot.keyPress(KeyEvent.VK_ENTER);
		robot.keyRelease(KeyEvent.VK_ENTER);	
		Thread.sleep(SleepBtwClicks);
		robot.keyPress(KeyEvent.VK_ENTER);
		robot.keyRelease(KeyEvent.VK_ENTER);
		
	    //Move Cursor to select "Power User" tab
		Thread.sleep(NewPageLoad);
		Thread.sleep(NewPageLoad);
		Thread.sleep(NewPageLoad);
		Thread.sleep(NewPageLoad);
		MAX_X = 285;
		MAX_Y = 148;
	    robot.mouseMove(MAX_X, MAX_Y);
	    Thread.sleep(SleepBtwClicks);
	    robot.mousePress(InputEvent.BUTTON1_MASK);
	    robot.mouseRelease(InputEvent.BUTTON1_MASK);
	    
			   
	    //Click on enter code "LN005"
		Thread.sleep(NewPageLoad);
		MAX_X = 52;
		MAX_Y = 195;
	    robot.mouseMove(MAX_X, MAX_Y);
	    Thread.sleep(SleepBtwClicks);
	    robot.mousePress(InputEvent.BUTTON1_MASK);
	    robot.mouseRelease(InputEvent.BUTTON1_MASK);
	    
	  //Enter code "LN005"
		Thread.sleep(SleepBtwClicks);
		robot.keyPress(KeyEvent.VK_L);
		robot.keyRelease(KeyEvent.VK_L);
		robot.keyPress(KeyEvent.VK_N);
		robot.keyRelease(KeyEvent.VK_N);
		robot.keyPress(KeyEvent.VK_0);
		robot.keyRelease(KeyEvent.VK_0);
		robot.keyPress(KeyEvent.VK_0);
		robot.keyRelease(KeyEvent.VK_0);
		robot.keyPress(KeyEvent.VK_5);
		robot.keyRelease(KeyEvent.VK_5);

		
		//Click on enter code "LN115"
		Thread.sleep(SleepBtwClicks);
		MAX_X = 52;
		MAX_Y = 215;
	    robot.mouseMove(MAX_X, MAX_Y);
	    Thread.sleep(SleepBtwClicks);
	    robot.mousePress(InputEvent.BUTTON1_MASK);
	    robot.mouseRelease(InputEvent.BUTTON1_MASK);
		Thread.sleep(SleepBtwClicks);
		robot.keyPress(KeyEvent.VK_ENTER);
		robot.keyRelease(KeyEvent.VK_ENTER);
	    
	  //Enter code "LN115"
		Thread.sleep(SleepBtwClicks);
		robot.keyPress(KeyEvent.VK_L);
		robot.keyRelease(KeyEvent.VK_L);
		robot.keyPress(KeyEvent.VK_N);
		robot.keyRelease(KeyEvent.VK_N);
		robot.keyPress(KeyEvent.VK_1);
		robot.keyRelease(KeyEvent.VK_1);
		robot.keyPress(KeyEvent.VK_1);
		robot.keyRelease(KeyEvent.VK_1);
		robot.keyPress(KeyEvent.VK_5);
		robot.keyRelease(KeyEvent.VK_5);
		
//		//Click on enter code "DT051"
//				Thread.sleep(SleepBtwClicks);
//				MAX_X = 52;
//				MAX_Y = 235;
//			    robot.mouseMove(MAX_X, MAX_Y);
//			    Thread.sleep(SleepBtwClicks);
//			    robot.mousePress(InputEvent.BUTTON1_MASK);
//			    robot.mouseRelease(InputEvent.BUTTON1_MASK);
//				Thread.sleep(SleepBtwClicks);
//				robot.keyPress(KeyEvent.VK_ENTER);
//				robot.keyRelease(KeyEvent.VK_ENTER);
//			    
//			  //Enter code "DT051"
//				Thread.sleep(SleepBtwClicks);
//				robot.keyPress(KeyEvent.VK_D);
//				robot.keyRelease(KeyEvent.VK_D);
//				robot.keyPress(KeyEvent.VK_T);
//				robot.keyRelease(KeyEvent.VK_T);
//				robot.keyPress(KeyEvent.VK_0);
//				robot.keyRelease(KeyEvent.VK_0);
//				robot.keyPress(KeyEvent.VK_5);
//				robot.keyRelease(KeyEvent.VK_5);
//				robot.keyPress(KeyEvent.VK_1);
//				robot.keyRelease(KeyEvent.VK_1);
		
				//Click on enter code "LN030"
				Thread.sleep(SleepBtwClicks);
				MAX_X = 52;
				MAX_Y = 235;
			    robot.mouseMove(MAX_X, MAX_Y);
			    Thread.sleep(SleepBtwClicks);
			    robot.mousePress(InputEvent.BUTTON1_MASK);
			    robot.mouseRelease(InputEvent.BUTTON1_MASK);
				Thread.sleep(SleepBtwClicks);
				robot.keyPress(KeyEvent.VK_ENTER);
				robot.keyRelease(KeyEvent.VK_ENTER);
			    
			  //Enter code "LN030"
				Thread.sleep(SleepBtwClicks);
				robot.keyPress(KeyEvent.VK_L);
				robot.keyRelease(KeyEvent.VK_L);
				robot.keyPress(KeyEvent.VK_N);
				robot.keyRelease(KeyEvent.VK_N);
				robot.keyPress(KeyEvent.VK_0);
				robot.keyRelease(KeyEvent.VK_0);
				robot.keyPress(KeyEvent.VK_3);
				robot.keyRelease(KeyEvent.VK_3);
				robot.keyPress(KeyEvent.VK_0);
				robot.keyRelease(KeyEvent.VK_0);
				
				
				//Click on enter code "LA700"
				Thread.sleep(SleepBtwClicks);
				MAX_X = 52;
				MAX_Y = 255;
			    robot.mouseMove(MAX_X, MAX_Y);
			    Thread.sleep(SleepBtwClicks);
			    robot.mousePress(InputEvent.BUTTON1_MASK);
			    robot.mouseRelease(InputEvent.BUTTON1_MASK);
				Thread.sleep(SleepBtwClicks);
				robot.keyPress(KeyEvent.VK_ENTER);
				robot.keyRelease(KeyEvent.VK_ENTER);
				//There is extra pop-up after LN030
				Thread.sleep(SleepBtwClicks);
				robot.keyPress(KeyEvent.VK_ENTER);
				robot.keyRelease(KeyEvent.VK_ENTER);
			    
			  //Enter code "LA700"
				Thread.sleep(SleepBtwClicks);
				robot.keyPress(KeyEvent.VK_L);
				robot.keyRelease(KeyEvent.VK_L);
				robot.keyPress(KeyEvent.VK_A);
				robot.keyRelease(KeyEvent.VK_A);
				robot.keyPress(KeyEvent.VK_7);
				robot.keyRelease(KeyEvent.VK_7);
				robot.keyPress(KeyEvent.VK_0);
				robot.keyRelease(KeyEvent.VK_0);
				robot.keyPress(KeyEvent.VK_0);
				robot.keyRelease(KeyEvent.VK_0);
								
				//Click on enter code "LC705"
				Thread.sleep(SleepBtwClicks);
				MAX_X = 52;
				MAX_Y = 275;
			    robot.mouseMove(MAX_X, MAX_Y);
			    Thread.sleep(SleepBtwClicks);
			    robot.mousePress(InputEvent.BUTTON1_MASK);
			    robot.mouseRelease(InputEvent.BUTTON1_MASK);
				Thread.sleep(SleepBtwClicks);
				robot.keyPress(KeyEvent.VK_ENTER);
				robot.keyRelease(KeyEvent.VK_ENTER);
			    
			  //Enter code "LC705"
				Thread.sleep(SleepBtwClicks);
				robot.keyPress(KeyEvent.VK_L);
				robot.keyRelease(KeyEvent.VK_L);
				robot.keyPress(KeyEvent.VK_C);
				robot.keyRelease(KeyEvent.VK_C);
				robot.keyPress(KeyEvent.VK_7);
				robot.keyRelease(KeyEvent.VK_7);
				robot.keyPress(KeyEvent.VK_0);
				robot.keyRelease(KeyEvent.VK_0);
				robot.keyPress(KeyEvent.VK_5);
				robot.keyRelease(KeyEvent.VK_5);
		
				//Click on enter code "LI750"
				Thread.sleep(SleepBtwClicks);
				MAX_X = 52;
				MAX_Y = 295;
			    robot.mouseMove(MAX_X, MAX_Y);
			    Thread.sleep(SleepBtwClicks);
			    robot.mousePress(InputEvent.BUTTON1_MASK);
			    robot.mouseRelease(InputEvent.BUTTON1_MASK);
				Thread.sleep(SleepBtwClicks);
				robot.keyPress(KeyEvent.VK_ENTER);
				robot.keyRelease(KeyEvent.VK_ENTER);
			    
			  //Enter code "LI750"
				Thread.sleep(SleepBtwClicks);
				robot.keyPress(KeyEvent.VK_L);
				robot.keyRelease(KeyEvent.VK_L);
				robot.keyPress(KeyEvent.VK_I);
				robot.keyRelease(KeyEvent.VK_I);
				robot.keyPress(KeyEvent.VK_7);
				robot.keyRelease(KeyEvent.VK_7);
				robot.keyPress(KeyEvent.VK_5);
				robot.keyRelease(KeyEvent.VK_5);
				robot.keyPress(KeyEvent.VK_0);
				robot.keyRelease(KeyEvent.VK_0);

				//Click on enter code "LA790"
				Thread.sleep(SleepBtwClicks);
				MAX_X = 52;
				MAX_Y = 315;
			    robot.mouseMove(MAX_X, MAX_Y);
			    Thread.sleep(SleepBtwClicks);
			    robot.mousePress(InputEvent.BUTTON1_MASK);
			    robot.mouseRelease(InputEvent.BUTTON1_MASK);
				Thread.sleep(SleepBtwClicks);
				robot.keyPress(KeyEvent.VK_ENTER);
				robot.keyRelease(KeyEvent.VK_ENTER);
			    
			  //Enter code "LA790"
				Thread.sleep(SleepBtwClicks);
				robot.keyPress(KeyEvent.VK_L);
				robot.keyRelease(KeyEvent.VK_L);
				robot.keyPress(KeyEvent.VK_A);
				robot.keyRelease(KeyEvent.VK_A);
				robot.keyPress(KeyEvent.VK_7);
				robot.keyRelease(KeyEvent.VK_7);
				robot.keyPress(KeyEvent.VK_9);
				robot.keyRelease(KeyEvent.VK_9);
				robot.keyPress(KeyEvent.VK_0);
				robot.keyRelease(KeyEvent.VK_0);				
		

				//Click on enter code "LA045"
				Thread.sleep(SleepBtwClicks);
				MAX_X = 52;
				MAX_Y = 335;
			    robot.mouseMove(MAX_X, MAX_Y);
			    Thread.sleep(SleepBtwClicks);
			    robot.mousePress(InputEvent.BUTTON1_MASK);
			    robot.mouseRelease(InputEvent.BUTTON1_MASK);
				Thread.sleep(SleepBtwClicks);
				robot.keyPress(KeyEvent.VK_ENTER);
				robot.keyRelease(KeyEvent.VK_ENTER);
			    
			  //Enter code "LA045"
				Thread.sleep(SleepBtwClicks);
				robot.keyPress(KeyEvent.VK_L);
				robot.keyRelease(KeyEvent.VK_L);
				robot.keyPress(KeyEvent.VK_A);
				robot.keyRelease(KeyEvent.VK_A);
				robot.keyPress(KeyEvent.VK_0);
				robot.keyRelease(KeyEvent.VK_0);
				robot.keyPress(KeyEvent.VK_4);
				robot.keyRelease(KeyEvent.VK_4);
				robot.keyPress(KeyEvent.VK_5);
				robot.keyRelease(KeyEvent.VK_5);
				
				Thread.sleep(SleepBtwClicks);
				MAX_X = 52;
				MAX_Y = 355;
			    robot.mouseMove(MAX_X, MAX_Y);
			    Thread.sleep(SleepBtwClicks);
			    
			    robot.mousePress(InputEvent.BUTTON1_MASK);
			    robot.mouseRelease(InputEvent.BUTTON1_MASK);
				Thread.sleep(SleepBtwClicks);
				robot.keyPress(KeyEvent.VK_ENTER);
				robot.keyRelease(KeyEvent.VK_ENTER);
				
				Thread.sleep(2000);
				//Click on avilable servivce
				
				Thread.sleep(SleepBtwClicks);
				MAX_X = 70;
				MAX_Y = 144;
			    robot.mouseMove(MAX_X, MAX_Y);
			    Thread.sleep(SleepBtwClicks);
			    robot.mousePress(InputEvent.BUTTON1_MASK);
			    robot.mouseRelease(InputEvent.BUTTON1_MASK);
				Thread.sleep(SleepBtwClicks);
					
				
				Thread.sleep(4000);
				// Filter
				Thread.sleep(SleepBtwClicks);
				MAX_X = 1023;
				MAX_Y = 341;
			    robot.mouseMove(MAX_X, MAX_Y);
			    Thread.sleep(SleepBtwClicks);
			    robot.mousePress(InputEvent.BUTTON1_MASK);
			    robot.mouseRelease(InputEvent.BUTTON1_MASK);
				Thread.sleep(SleepBtwClicks);
					
				
				Thread.sleep(2000);
				Thread.sleep(SleepBtwClicks);
				MAX_X = 1023;
				MAX_Y = 357;
			    robot.mouseMove(MAX_X, MAX_Y);
			    Thread.sleep(SleepBtwClicks);
			    robot.mousePress(InputEvent.BUTTON1_MASK);
			    robot.mouseRelease(InputEvent.BUTTON1_MASK);
				Thread.sleep(SleepBtwClicks);
			
				Thread.sleep(2000);
				// select tele
				
				Thread.sleep(SleepBtwClicks);
				MAX_X = 44;
				MAX_Y = 193;
			    robot.mouseMove(MAX_X, MAX_Y);
			    Thread.sleep(SleepBtwClicks);
			    robot.mousePress(InputEvent.BUTTON1_MASK);
			    robot.mouseRelease(InputEvent.BUTTON1_MASK);
				Thread.sleep(SleepBtwClicks);
				
				
				Thread.sleep(2000);
				//selct HP active
				
				Thread.sleep(SleepBtwClicks);
				MAX_X = 1042;
				MAX_Y = 273;
			    robot.mouseMove(MAX_X, MAX_Y);
			    Thread.sleep(SleepBtwClicks);
			    robot.mousePress(InputEvent.BUTTON1_MASK);
			    robot.mouseRelease(InputEvent.BUTTON1_MASK);
				Thread.sleep(SleepBtwClicks);
				
				//drop down
				Thread.sleep(2000);
				Thread.sleep(SleepBtwClicks);
				MAX_X = 1329;
				MAX_Y = 290;
			    robot.mouseMove(MAX_X, MAX_Y);
			    Thread.sleep(SleepBtwClicks);
			    robot.mousePress(InputEvent.BUTTON1_MASK);
			    robot.mouseRelease(InputEvent.BUTTON1_MASK);
				Thread.sleep(SleepBtwClicks);
			
				
				
				Thread.sleep(2000);
				//Select HP Auto
				Thread.sleep(SleepBtwClicks);
				MAX_X = 1042;
				MAX_Y = 273;
			    robot.mouseMove(MAX_X, MAX_Y);
			    Thread.sleep(SleepBtwClicks);
			    robot.mousePress(InputEvent.BUTTON1_MASK);
			    robot.mouseRelease(InputEvent.BUTTON1_MASK);
				Thread.sleep(SleepBtwClicks);
				
				//drop down
				Thread.sleep(2000);
				Thread.sleep(SleepBtwClicks);
				MAX_X = 1329;
				MAX_Y = 290;
			    robot.mouseMove(MAX_X, MAX_Y);
			    Thread.sleep(SleepBtwClicks);
			    robot.mousePress(InputEvent.BUTTON1_MASK);
			    robot.mouseRelease(InputEvent.BUTTON1_MASK);
				Thread.sleep(SleepBtwClicks);
				
				Thread.sleep(2000);
				//Select cid
				Thread.sleep(SleepBtwClicks);
				MAX_X = 1042;
				MAX_Y = 273;
			    robot.mouseMove(MAX_X, MAX_Y);
			    Thread.sleep(SleepBtwClicks);
			    robot.mousePress(InputEvent.BUTTON1_MASK);
			    robot.mouseRelease(InputEvent.BUTTON1_MASK);
				Thread.sleep(SleepBtwClicks);
				
				Thread.sleep(2000);
				Thread.sleep(SleepBtwClicks);
				MAX_X = 1224;
				MAX_Y = 478;
			    robot.mouseMove(MAX_X, MAX_Y);
			    Thread.sleep(SleepBtwClicks);
			    robot.mousePress(InputEvent.BUTTON1_MASK);
			    robot.mouseRelease(InputEvent.BUTTON1_MASK);
				Thread.sleep(SleepBtwClicks);
				robot.keyPress(KeyEvent.VK_ENTER);
				robot.keyRelease(KeyEvent.VK_ENTER);
				
				Thread.sleep(2000);	
				Thread.sleep(SleepBtwClicks);
				robot.keyPress(KeyEvent.VK_TAB);
				robot.keyRelease(KeyEvent.VK_TAB);
				Thread.sleep(SleepBtwClicks);
				robot.keyPress(KeyEvent.VK_ENTER);
				robot.keyRelease(KeyEvent.VK_ENTER);
				Thread.sleep(SleepBtwClicks);
				robot.keyPress(KeyEvent.VK_ENTER);
				robot.keyRelease(KeyEvent.VK_ENTER);
				
				Thread.sleep(4000);	

									

		
	    //Next page Alt+N
	    Thread.sleep(SleepBtwClicks);
		robot.keyPress(KeyEvent.VK_ALT);
		robot.keyPress(KeyEvent.VK_N);
		robot.keyRelease(KeyEvent.VK_N);
		robot.keyRelease(KeyEvent.VK_ALT);
		Thread.sleep(SleepBtwClicks);
		robot.keyPress(KeyEvent.VK_ENTER);
		robot.keyRelease(KeyEvent.VK_ENTER);
		
	    //Enter email address
		Thread.sleep(NewPageLoad);
		Thread.sleep(NewPageLoad);
		MAX_X = 756;
		MAX_Y = 224;
	    robot.mouseMove(MAX_X, MAX_Y);
	    Thread.sleep(SleepBtwClicks);
	    robot.mousePress(InputEvent.BUTTON1_MASK);
	    robot.mouseRelease(InputEvent.BUTTON1_MASK);
	    Thread.sleep(SleepBtwClicks);
		robot.keyPress(KeyEvent.VK_X);
		robot.keyRelease(KeyEvent.VK_X);
		
	    //Next page Alt+N
	    Thread.sleep(SleepBtwClicks);
		robot.keyPress(KeyEvent.VK_ALT);
		robot.keyPress(KeyEvent.VK_N);
		robot.keyRelease(KeyEvent.VK_N);
		robot.keyRelease(KeyEvent.VK_ALT);
		
		
		
		//select provisioning
		
				Thread.sleep(NewPageLoad);
				Thread.sleep(NewPageLoad);
				MAX_X = 340;
				MAX_Y = 621;
			    robot.mouseMove(MAX_X, MAX_Y);
			    Thread.sleep(SleepBtwClicks);
			    robot.mousePress(InputEvent.BUTTON1_MASK);
			    robot.mouseRelease(InputEvent.BUTTON1_MASK);
			    Thread.sleep(SleepBtwClicks);
				robot.keyPress(KeyEvent.VK_0);
				robot.keyRelease(KeyEvent.VK_0);
				
				
				//Drop down
				
				Thread.sleep(SleepBtwClicks);
				MAX_X = 1322;
				MAX_Y = 525;
			    robot.mouseMove(MAX_X, MAX_Y);
			    Thread.sleep(SleepBtwClicks);
			    
			    for(int g=0; g<6; g++)
			    {
			    robot.mousePress(InputEvent.BUTTON1_MASK);
			    robot.mouseRelease(InputEvent.BUTTON1_MASK);
				Thread.sleep(SleepBtwClicks);
			    }
			   
			  
				Thread.sleep(NewPageLoad);
				MAX_X = 824;
				MAX_Y = 520;
			    robot.mouseMove(MAX_X, MAX_Y);
			    Thread.sleep(SleepBtwClicks);
			    robot.mousePress(InputEvent.BUTTON1_MASK);
			    robot.mouseRelease(InputEvent.BUTTON1_MASK);
			    Thread.sleep(SleepBtwClicks);
				robot.keyPress(KeyEvent.VK_Y);
				robot.keyRelease(KeyEvent.VK_Y);
				Thread.sleep(1000);
				
				
				 //Next page Alt+N
			    Thread.sleep(SleepBtwClicks);
				robot.keyPress(KeyEvent.VK_ALT);
				robot.keyPress(KeyEvent.VK_N);
				robot.keyRelease(KeyEvent.VK_N);
				robot.keyRelease(KeyEvent.VK_ALT);
			    
	    //Enter Job details
		Thread.sleep(NewPageLoad);
		robot.keyPress(KeyEvent.VK_TAB);
		robot.keyRelease(KeyEvent.VK_TAB);
		Thread.sleep(SleepBtwClicks);
		robot.keyPress(KeyEvent.VK_0);
		robot.keyRelease(KeyEvent.VK_0);
		Thread.sleep(SleepBtwClicks);
		robot.keyPress(KeyEvent.VK_9);
		robot.keyRelease(KeyEvent.VK_9);
		robot.keyPress(KeyEvent.VK_9);
		robot.keyRelease(KeyEvent.VK_9);
		robot.keyPress(KeyEvent.VK_6);
		robot.keyRelease(KeyEvent.VK_6);
		robot.keyPress(KeyEvent.VK_6);
		robot.keyRelease(KeyEvent.VK_6);
		Thread.sleep(SleepBtwClicks);
		robot.keyPress(KeyEvent.VK_TAB);
		robot.keyRelease(KeyEvent.VK_TAB);
		Thread.sleep(SleepBtwClicks);
		robot.keyPress(KeyEvent.VK_X);
		robot.keyRelease(KeyEvent.VK_X);
		robot.keyPress(KeyEvent.VK_T);
		robot.keyRelease(KeyEvent.VK_T);
		
		//Enter Equipment detail - 
		Thread.sleep(SleepBtwClicks);
		MAX_X = 41;
		MAX_Y = 433;
	    robot.mouseMove(MAX_X, MAX_Y);
	    Thread.sleep(SleepBtwClicks);
        robot.mousePress(InputEvent.BUTTON1_MASK);
        robot.mouseRelease(InputEvent.BUTTON1_MASK);
        Thread.sleep(SleepBtwClicks);
           robot.keyPress(KeyEvent.VK_ALT);
           robot.keyPress(KeyEvent.VK_A);
           robot.keyRelease(KeyEvent.VK_A);
           robot.keyRelease(KeyEvent.VK_ALT);
        Thread.sleep(SleepBtwClicks);
        Thread.sleep(SleepBtwClicks);
           robot.keyPress(KeyEvent.VK_SHIFT);
           robot.keyPress(KeyEvent.VK_8);
           robot.keyRelease(KeyEvent.VK_8);
           robot.keyRelease(KeyEvent.VK_SHIFT);
           Thread.sleep(SleepBtwClicks);
           robot.keyPress(KeyEvent.VK_TAB);
           robot.keyRelease(KeyEvent.VK_TAB);
        Thread.sleep(SleepBtwClicks);
           robot.keyPress(KeyEvent.VK_TAB);
           robot.keyRelease(KeyEvent.VK_TAB);
           Thread.sleep(SleepBtwClicks);
           robot.keyPress(KeyEvent.VK_N);
           robot.keyRelease(KeyEvent.VK_N);
          
        Thread.sleep(SleepBtwClicks);
        
        
        Thread.sleep(SleepBtwClicks);
           robot.keyPress(KeyEvent.VK_ALT);
           robot.keyPress(KeyEvent.VK_A);
           robot.keyRelease(KeyEvent.VK_A);
           robot.keyRelease(KeyEvent.VK_ALT);
        Thread.sleep(SleepBtwClicks);
        Thread.sleep(SleepBtwClicks);
           robot.keyPress(KeyEvent.VK_SHIFT);
           robot.keyPress(KeyEvent.VK_8);
           robot.keyRelease(KeyEvent.VK_8);
           robot.keyRelease(KeyEvent.VK_SHIFT);
           Thread.sleep(SleepBtwClicks);
           robot.keyPress(KeyEvent.VK_TAB);
           robot.keyRelease(KeyEvent.VK_TAB);
        Thread.sleep(SleepBtwClicks);
           robot.keyPress(KeyEvent.VK_TAB);
           robot.keyRelease(KeyEvent.VK_TAB);
           Thread.sleep(SleepBtwClicks);
           robot.keyPress(KeyEvent.VK_N);
           robot.keyRelease(KeyEvent.VK_N);
           
           Thread.sleep(SleepBtwClicks);
           robot.keyPress(KeyEvent.VK_ALT);
           robot.keyPress(KeyEvent.VK_A);
           robot.keyRelease(KeyEvent.VK_A);
           robot.keyRelease(KeyEvent.VK_ALT);
        Thread.sleep(SleepBtwClicks);
        Thread.sleep(SleepBtwClicks);
           robot.keyPress(KeyEvent.VK_SHIFT);
           robot.keyPress(KeyEvent.VK_8);
           robot.keyRelease(KeyEvent.VK_8);
           robot.keyRelease(KeyEvent.VK_SHIFT);
           Thread.sleep(SleepBtwClicks);
           robot.keyPress(KeyEvent.VK_TAB);
           robot.keyRelease(KeyEvent.VK_TAB);
        Thread.sleep(SleepBtwClicks);
           robot.keyPress(KeyEvent.VK_TAB);
           robot.keyRelease(KeyEvent.VK_TAB);
           Thread.sleep(SleepBtwClicks);
           robot.keyPress(KeyEvent.VK_R);
           robot.keyRelease(KeyEvent.VK_R);
           robot.keyPress(KeyEvent.VK_P);
           robot.keyRelease(KeyEvent.VK_P);
          
       
	
		
		
		//Next Page
		Thread.sleep(SleepBtwClicks);
		robot.keyPress(KeyEvent.VK_ALT);
		robot.keyPress(KeyEvent.VK_N);
		robot.keyRelease(KeyEvent.VK_N);
		robot.keyRelease(KeyEvent.VK_ALT);

		  //*Validate the page to proceed - Finish page
  		ImageMatched = false;
  		while (!ImageMatched){
  //*Save the current image at "Search" button position
  	       try {
  	          /*
  	           * Let the program wait for 5 seconds to allow you to open the
  	           * window whose screenshot has to be captured
  	           */
  	           Thread.sleep(NewPageLoad);
  	           String fileName = "C://Driver//Finish_Current.jpg";
  	 
  	           // Define an area of size 500*400 starting at coordinates (10,50)
  	           //Rectangle rectArea = new Rectangle(24, 218, 56, 12);
  	           Rectangle rectArea = new Rectangle(1182, 678, 79, 20);
  	           BufferedImage screenFullImage = robot.createScreenCapture(rectArea);
  	           ImageIO.write(screenFullImage, "jpg", new File(fileName));
  	 

  	       } catch (IOException | InterruptedException ex) {
  	                System.err.println(ex);
  	       }
  		
  //Compare the image with saved image - 
  	       BufferedImage imgA = null; 
  	       BufferedImage imgB = null; 

  	       try
  	       { 
  	           File fileA = new File("C://Driver//Finish_Current.jpg"); 
  	           File fileB = new File("C://Driver//Finish.jpg"); 

  	           imgA = ImageIO.read(fileA); 
  	           imgB = ImageIO.read(fileB); 
  	       } 
  	       catch (IOException e) 
  	       { 
  	           System.out.println(e); 
  	       } 
  	       int width1 = imgA.getWidth(); 
  	       int width2 = imgB.getWidth(); 
  	       int height1 = imgA.getHeight(); 
  	       int height2 = imgB.getHeight(); 

  	     if ((width1 == width2) && (height1 == height2)) 
	       {
	    	   long difference = 0; 
	           for (int y = 0; y < height1; y++) 
	           { 
	               for (int x = 0; x < width1; x++) 
	               { 
	                   int rgbA = imgA.getRGB(x, y); 
	                   int rgbB = imgB.getRGB(x, y); 
	                   int redA = (rgbA >> 16) & 0xff; 
	                   int greenA = (rgbA >> 8) & 0xff; 
	                   int blueA = (rgbA) & 0xff; 
	                   int redB = (rgbB >> 16) & 0xff; 
	                   int greenB = (rgbB >> 8) & 0xff; 
	                   int blueB = (rgbB) & 0xff; 
	                   difference += Math.abs(redA - redB); 
	                   difference += Math.abs(greenA - greenB); 
	                   difference += Math.abs(blueA - blueB); 
			           
	               } 
	           }
	           System.out.println("difference is"+difference);
	           if (difference <= 8000){
	        	   System.out.println("Image matched");
	        	   ImageMatched = true;
	           }
	       System.out.println("first BP2");
	       }
	       else
	       { 
	    	  // ImageMatched = false;
	    	   Thread.sleep(NewPageLoad);
	    	   System.out.println("first BP3");
	       }
		}   
		
		//Enter to Finish
		Thread.sleep(SleepBtwClicks);
		Thread.sleep(NewPageLoad);
		Thread.sleep(NewPageLoad);
		robot.keyPress(KeyEvent.VK_ENTER);
		robot.keyRelease(KeyEvent.VK_ENTER);
		

		
		/*
	    //Copy Account Number and write in spreadsheet
		Thread.sleep(NewPageLoad);
		Thread.sleep(SleepBtwClicks);
		MAX_X = 1058;
		MAX_Y = 146;
	    robot.mouseMove(MAX_X, MAX_Y);
	    robot.mousePress(InputEvent.BUTTON1_MASK);
	    robot.mouseRelease(InputEvent.BUTTON1_MASK);
	    robot.mousePress(InputEvent.BUTTON1_MASK);
	    robot.mouseRelease(InputEvent.BUTTON1_MASK);
		Thread.sleep(SleepBtwClicks);
	    robot.keyPress(KeyEvent.VK_CONTROL);
	    robot.keyPress(KeyEvent.VK_C);
	    robot.keyRelease(KeyEvent.VK_C);
		robot.keyRelease(KeyEvent.VK_CONTROL);
	       
		//Copy the AccountNumber from Clipboard to String
		Thread.sleep(SleepBtwClicks);
		Toolkit toolkit = Toolkit.getDefaultToolkit();
		Clipboard clipboard = toolkit.getSystemClipboard();
		String AccountNumber = (String) clipboard.getData(DataFlavor.stringFlavor);
		 
		
		//Save the AccountNumber in spreadsheet
		Cell cell=s.getRow(row).createCell(s.getRow(row).getPhysicalNumberOfCells());
		cell.setCellType(cell.CELL_TYPE_STRING);
		cell.setCellValue(AccountNumber);
		Thread.sleep(SleepBtwClicks);
		FileOutputStream fo=new FileOutputStream("C:\\Driver\\AutomatedLocationCreation.xlsx");
		wb.write(fo);
		
		//Copy Order ID and write in spreadsheet
		Thread.sleep(SleepBtwClicks);
		MAX_X = 809;
		MAX_Y = 169;
		robot.mouseMove(MAX_X, MAX_Y);
		robot.mousePress(InputEvent.BUTTON1_MASK);
		robot.mouseRelease(InputEvent.BUTTON1_MASK);
		robot.mousePress(InputEvent.BUTTON1_MASK);
		robot.mouseRelease(InputEvent.BUTTON1_MASK);
		Thread.sleep(SleepBtwClicks);
		robot.keyPress(KeyEvent.VK_CONTROL);
		robot.keyPress(KeyEvent.VK_C);
		robot.keyRelease(KeyEvent.VK_C);
		robot.keyRelease(KeyEvent.VK_CONTROL);
			       
		//Copy the AccountNumber from Clipboard to String
		Thread.sleep(SleepBtwClicks);
		toolkit = Toolkit.getDefaultToolkit();
		clipboard = toolkit.getSystemClipboard();
		OrderID = (String) clipboard.getData(DataFlavor.stringFlavor);
		 
			
		//Save the AccountNumber in spreadsheet
		Cell cell6=s.getRow(row).createCell(s.getRow(row).getPhysicalNumberOfCells());
		cell6.setCellType(cell6.CELL_TYPE_STRING);
		cell6.setCellValue(OrderID);
		Thread.sleep(SleepBtwClicks);
		wb.write(fo);
		
	    */
	
		}
			catch(Exception e){
				
				Cell cell=s.getRow(row).createCell(s.getRow(row).getPhysicalNumberOfCells());
				cell.setCellType(cell.CELL_TYPE_STRING);
				cell.setCellValue(e.toString());
				FileOutputStream fo=new FileOutputStream("C:\\Driver\\AutomatedAccountCreation.xlsx");
				wb.write(fo);
					
						}
		}
	    }
	    
	   
	}
	
