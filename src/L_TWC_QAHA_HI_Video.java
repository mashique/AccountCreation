

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



	public class L_TWC_QAHA_HI_Video {
    
	    public static final int SleepBtwClicks = 500;
	    public static final int NewPageLoad = 4000;
	    public static int MAX_Y = 440;
	    public static int MAX_X = 110;	
	    public static int StartExecution = 2000;
	    public static String FirstName;
	    public static String LocationID;
	    public static String AccountNumber;
	    public static String OrderID;
	    static boolean ImageMatched;
	    
	    public static void main(String... args) throws Exception {
	    Robot robot = new Robot();

		FileInputStream f=new FileInputStream("D:\\Softwares\\QakaAccountCreation.xlsx");
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

		
	    //Move Cursor to enter LocationID
		MAX_X = 774;
		MAX_Y = 407;
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
		MAX_X = 72;
		MAX_Y = 186;
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
		Thread.sleep(NewPageLoad);
		
		//Click on enter code "LN115"
		Thread.sleep(SleepBtwClicks);
		MAX_X = 72;
		MAX_Y = 202;
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
		Thread.sleep(NewPageLoad);


				//Click on enter code "LA045"
				Thread.sleep(SleepBtwClicks);
				MAX_X = 72;
				MAX_Y = 220;
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
				Thread.sleep(NewPageLoad);
				
					
		//Click on next row
		Thread.sleep(SleepBtwClicks);
		MAX_X = 72;
		MAX_Y = 235;
	    robot.mouseMove(MAX_X, MAX_Y);
	    Thread.sleep(SleepBtwClicks);
	    robot.mousePress(InputEvent.BUTTON1_MASK);
	    robot.mouseRelease(InputEvent.BUTTON1_MASK);
		Thread.sleep(SleepBtwClicks);
		robot.keyPress(KeyEvent.VK_ENTER);
		robot.keyRelease(KeyEvent.VK_ENTER);
		Thread.sleep(SleepBtwClicks);
		robot.keyPress(KeyEvent.VK_ENTER);
		robot.keyRelease(KeyEvent.VK_ENTER);	
	   
		
	    //Next page Alt+N
	    Thread.sleep(SleepBtwClicks);
		robot.keyPress(KeyEvent.VK_ALT);
		robot.keyPress(KeyEvent.VK_N);
		robot.keyRelease(KeyEvent.VK_N);
		robot.keyRelease(KeyEvent.VK_ALT);
		Thread.sleep(SleepBtwClicks);
		robot.keyPress(KeyEvent.VK_ENTER);
		robot.keyRelease(KeyEvent.VK_ENTER);
		Thread.sleep(SleepBtwClicks);
		robot.keyPress(KeyEvent.VK_ENTER);
		robot.keyRelease(KeyEvent.VK_ENTER);		

		Thread.sleep(NewPageLoad);
		MAX_X = 847;
		MAX_Y = 434;
	    robot.mouseMove(MAX_X, MAX_Y);
	    Thread.sleep(SleepBtwClicks);
	    robot.mousePress(InputEvent.BUTTON1_MASK);
	    robot.mouseRelease(InputEvent.BUTTON1_MASK);	
	    Thread.sleep(SleepBtwClicks);
	    robot.mousePress(InputEvent.BUTTON1_MASK);
	    robot.mouseRelease(InputEvent.BUTTON1_MASK);
	    
	    
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
		
		
		//Select VIP code
		Thread.sleep(NewPageLoad);
		MAX_X = 800;
		MAX_Y = 587;
	    robot.mouseMove(MAX_X, MAX_Y);
	    Thread.sleep(SleepBtwClicks);
	    robot.mousePress(InputEvent.BUTTON1_MASK);
	    robot.mouseRelease(InputEvent.BUTTON1_MASK);
	    Thread.sleep(SleepBtwClicks);
		robot.keyPress(KeyEvent.VK_I);
		robot.keyRelease(KeyEvent.VK_I);
		
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
		

		//Enter Equipment detail
		Thread.sleep(SleepBtwClicks);
		MAX_X = 47;	
		MAX_Y = 423;
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
	    Thread.sleep(SleepBtwClicks);
		
		
		//Next Page
		Thread.sleep(SleepBtwClicks);
		robot.keyPress(KeyEvent.VK_ALT);
		robot.keyPress(KeyEvent.VK_N);
		robot.keyRelease(KeyEvent.VK_N);
		robot.keyRelease(KeyEvent.VK_ALT);


		//Enter to Finish
		Thread.sleep(SleepBtwClicks);
		Thread.sleep(NewPageLoad);
		robot.keyPress(KeyEvent.VK_ENTER);
		robot.keyRelease(KeyEvent.VK_ENTER);
			

		

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
		FileOutputStream fo=new FileOutputStream("D:\\Softwares\\QakaAccountCreation.xlsx");
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
		Cell cell3=s.getRow(row).createCell(s.getRow(row).getPhysicalNumberOfCells());
		cell3.setCellType(cell3.CELL_TYPE_STRING);
		cell3.setCellValue(OrderID);
		Thread.sleep(SleepBtwClicks);
		wb.write(fo);

	    
	
		}
			catch(Exception e){
				
				Cell cell=s.getRow(row).createCell(s.getRow(row).getPhysicalNumberOfCells());
				cell.setCellType(cell.CELL_TYPE_STRING);
				cell.setCellValue(e.toString());
				FileOutputStream fo=new FileOutputStream("D:\\Softwares\\QakaAccountCreation.xlsx");
				wb.write(fo);
					
						}
		}
	    }
	    
	   
	}
	
