
import java.io.FileInputStream;
import java.awt.AWTException;
import java.awt.Frame;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.event.InputEvent;
import java.awt.event.KeyEvent;
import java.util.Random;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.StringSelection;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

	class AccountCreation_BHN {
	    public static final int FIVE_SECONDS = 1000;
	    public static final int SleepBtwClicks = 1500;
	    public static final int WaitForInstallPage = 5000;
	    public static final int WaitForPackagePage = 5000;
	    public static final int PageLoad = 5000;
	    public static final int SmallWait = 500;
	  //public static final int MAX_Y = 5;
	  //public static final int MAX_X = 5;
	    public static int MAX_Y = 440;
	    public static int MAX_X = 110;
	    public static int StartExecution = 2000;
	    public static int count = 0;
	    public static String theString = "123445";
	    public static String AccountNum;
	    public static String HouseNumber;
	    public static String FirstName;
	    public static String LastName;
	    public static String ServiceCodeSet1 = "LA700\nLA790\nLC707\nLI750\nLV750";
	    public static String ServiceQTY1 = "1\n1\n1\n1\n1";
	    public static String ServiceCodeSet2 = "LG015\nLG020\nLG025\nLG030\nLG035\nLZ020\nLZ100";
	    public static String ServiceQTY2 = "1\n1\n1\n1\n1\n1\n1";
	    public static String ServiceCodeSet3 = "LZ105\nLZ107";
	    public static String ServiceQTY3 = "1\n1";
	    
	    public static void main(String... args) throws Exception {
		Robot robot = new Robot();
		//Random random = new Random();
		//MAX_X  = OFFSET; 
		//MAX_Y = OFFSET;
		FileInputStream f=new FileInputStream("D:\\Softwares\\BHNAccount.xlsx");
		XSSFWorkbook wb = new XSSFWorkbook(f);
		XSSFSheet s= wb.getSheet("Sheet1");
		int totalRow=s.getPhysicalNumberOfRows();
		System.out.println("number of rows in spreadsheet"+totalRow);
		
		for(int row=1; row<totalRow; row++){
			Cell cell1=s.getRow(row).getCell(0);
			Cell cell2=s.getRow(row).getCell(1);
			Cell cell3=s.getRow(row).getCell(2);
			Cell cell4=s.getRow(row).getCell(3);
			AccountNum=cell1.getStringCellValue();
			HouseNumber=cell2.getStringCellValue();
			FirstName=cell3.getStringCellValue();
			LastName=cell4.getStringCellValue();

			

	    //First FIELD
	    //Enter Account number
		Thread.sleep(SleepBtwClicks);
		MAX_X = 211;
		MAX_Y = 85;
	    robot.mouseMove(MAX_X, MAX_Y);
	  
	    Thread.sleep(SleepBtwClicks);
	    robot.mousePress(InputEvent.BUTTON1_MASK);
	    robot.mouseRelease(InputEvent.BUTTON1_MASK);
	    robot.mousePress(InputEvent.BUTTON1_MASK);
	    robot.mouseRelease(InputEvent.BUTTON1_MASK);
	    
		StringSelection selection3 = new StringSelection(AccountNum);
		Clipboard clipboard3 = Toolkit.getDefaultToolkit().getSystemClipboard();
		clipboard3.setContents(selection3, selection3);
	    
		Thread.sleep(SleepBtwClicks);
	      robot.keyPress(KeyEvent.VK_CONTROL);
	      robot.keyPress(KeyEvent.VK_V);
	      robot.keyRelease(KeyEvent.VK_V);
	      robot.keyRelease(KeyEvent.VK_CONTROL);
	      Thread.sleep(SleepBtwClicks);
	      
	     //enter second part of account number-- HouseNumber
		MAX_X = 282;
		MAX_Y = 82;
		robot.mouseMove(MAX_X, MAX_Y);
		 
		   robot.mousePress(InputEvent.BUTTON1_MASK);
		   robot.mouseRelease(InputEvent.BUTTON1_MASK);
		   robot.mousePress(InputEvent.BUTTON1_MASK);
		   robot.mouseRelease(InputEvent.BUTTON1_MASK);
		   
		   System.out.println("value of housenumber is"+HouseNumber);
		
		selection3 = new StringSelection(HouseNumber);
		clipboard3 = Toolkit.getDefaultToolkit().getSystemClipboard();
		clipboard3.setContents(selection3, selection3);
		    
		Thread.sleep(SleepBtwClicks);
		     robot.keyPress(KeyEvent.VK_CONTROL);
		     robot.keyPress(KeyEvent.VK_V);
		     robot.keyRelease(KeyEvent.VK_V);
		     robot.keyRelease(KeyEvent.VK_CONTROL);
		Thread.sleep(SleepBtwClicks);   
		      
		   //enter Task
		MAX_X = 392;
			MAX_Y = 81;
		    robot.mouseMove(MAX_X, MAX_Y);
		  
		    robot.mousePress(InputEvent.BUTTON1_MASK);
		    robot.mouseRelease(InputEvent.BUTTON1_MASK);
		    robot.mousePress(InputEvent.BUTTON1_MASK);
		    robot.mouseRelease(InputEvent.BUTTON1_MASK);
		    
			selection3 = new StringSelection("02");
			clipboard3 = Toolkit.getDefaultToolkit().getSystemClipboard();
			clipboard3.setContents(selection3, selection3);
		    
			Thread.sleep(SleepBtwClicks);
		      robot.keyPress(KeyEvent.VK_CONTROL);
		      robot.keyPress(KeyEvent.VK_V);
		      robot.keyRelease(KeyEvent.VK_V);
		      robot.keyRelease(KeyEvent.VK_CONTROL);
	     
	      //press Enter Twice
	      robot.keyPress(KeyEvent.VK_ENTER);
	      robot.keyRelease(KeyEvent.VK_ENTER);
	      robot.keyPress(KeyEvent.VK_ENTER);
	      robot.keyRelease(KeyEvent.VK_ENTER);
	    
	    //click on install
	    Thread.sleep(WaitForInstallPage); 
		MAX_X = 579;
		MAX_Y = 374;
	    robot.mouseMove(MAX_X, MAX_Y);
	    robot.mousePress(InputEvent.BUTTON1_MASK);
	    robot.mouseRelease(InputEvent.BUTTON1_MASK);
	    robot.mousePress(InputEvent.BUTTON1_MASK);
	    robot.mouseRelease(InputEvent.BUTTON1_MASK);
	    
	    //double click on F15
	    Thread.sleep(WaitForInstallPage); 
		MAX_X = 579;
		MAX_Y = 358;
	    robot.mouseMove(MAX_X, MAX_Y);
	    robot.mousePress(InputEvent.BUTTON1_MASK);
	    robot.mouseRelease(InputEvent.BUTTON1_MASK);
	    robot.mousePress(InputEvent.BUTTON1_MASK);
	    robot.mouseRelease(InputEvent.BUTTON1_MASK);	    
	    
	   
	    
	   
	    //THIRD FIELD
	    //Enter First Name
	    Thread.sleep(PageLoad);
	    theString = FirstName;
		StringSelection selection4 = new StringSelection(theString);
		Clipboard clipboard4 = Toolkit.getDefaultToolkit().getSystemClipboard();
		clipboard4.setContents(selection4, selection4);
		MAX_X = 246;
		MAX_Y = 232;
	    robot.mouseMove(MAX_X, MAX_Y);
	    robot.mousePress(InputEvent.BUTTON1_MASK);
	    robot.mouseRelease(InputEvent.BUTTON1_MASK);
	    
	    Thread.sleep(SleepBtwClicks);
	      robot.keyPress(KeyEvent.VK_CONTROL);
	      robot.keyPress(KeyEvent.VK_V);
	      robot.keyRelease(KeyEvent.VK_V);
	      robot.keyRelease(KeyEvent.VK_CONTROL);

	    //Fourth FIELD
	    //Enter Last name
	    Thread.sleep(SleepBtwClicks);
	    theString = LastName;
		StringSelection selection5 = new StringSelection(theString);
		Clipboard clipboard5 = Toolkit.getDefaultToolkit().getSystemClipboard();
		clipboard5.setContents(selection5, selection5);
	    //MOVE MOUSE AND PASTE
		MAX_X = 650;
		MAX_Y = 235;
	    robot.mouseMove(MAX_X, MAX_Y);
	    Thread.sleep(SleepBtwClicks);
	    robot.mousePress(InputEvent.BUTTON1_MASK);
	    robot.mouseRelease(InputEvent.BUTTON1_MASK);
	    
	    Thread.sleep(SleepBtwClicks);
	      robot.keyPress(KeyEvent.VK_CONTROL);
	      robot.keyPress(KeyEvent.VK_V);
	      robot.keyRelease(KeyEvent.VK_V);
	      robot.keyRelease(KeyEvent.VK_CONTROL);
	    
	    //Press Enter to continue
	      Thread.sleep(SmallWait);
	      robot.keyPress(KeyEvent.VK_ENTER);
	      robot.keyRelease(KeyEvent.VK_ENTER);
	      
	    
	   //Enter Customer Type
	      robot.keyPress(KeyEvent.VK_6);
	      robot.keyRelease(KeyEvent.VK_6);
		    
	   //Press Enter to continue
	     Thread.sleep(SmallWait);
	     robot.keyPress(KeyEvent.VK_ENTER);
	     robot.keyRelease(KeyEvent.VK_ENTER);
		      
		 //Enter Customer Type41
	      robot.keyPress(KeyEvent.VK_G);
	      robot.keyRelease(KeyEvent.VK_G);
	   	    
		 //Press Enter Twice to continue
		    Thread.sleep(SmallWait);
		    robot.keyPress(KeyEvent.VK_ENTER);
		    robot.keyRelease(KeyEvent.VK_ENTER);
		     
		    Thread.sleep(SmallWait);
		    robot.keyPress(KeyEvent.VK_ENTER);
		    robot.keyRelease(KeyEvent.VK_ENTER);
		      
		   //Enter mandatory fiend Authorized by
		    Thread.sleep(SleepBtwClicks);
		      theString = "test";
				StringSelection selection7 = new StringSelection(theString);
				Clipboard clipboard7 = Toolkit.getDefaultToolkit().getSystemClipboard();
				clipboard7.setContents(selection7, selection7);
			    
			   Thread.sleep(SleepBtwClicks);
			    robot.keyPress(KeyEvent.VK_CONTROL);
			    robot.keyPress(KeyEvent.VK_V);
			    robot.keyRelease(KeyEvent.VK_V);
			    robot.keyRelease(KeyEvent.VK_CONTROL);
			      
			 //Press Enter Thrice to continue
		    Thread.sleep(SleepBtwClicks);
		    robot.keyPress(KeyEvent.VK_ENTER);
		    robot.keyRelease(KeyEvent.VK_ENTER);
		    
		    Thread.sleep(SleepBtwClicks);
		    robot.keyPress(KeyEvent.VK_ENTER);
		    robot.keyRelease(KeyEvent.VK_ENTER);
				     
		    //double click on F15
		    Thread.sleep(WaitForInstallPage); 
			MAX_X = 249;
			MAX_Y = 296;
		    robot.mouseMove(MAX_X, MAX_Y);
		    robot.mousePress(InputEvent.BUTTON1_MASK);
		    robot.mouseRelease(InputEvent.BUTTON1_MASK);
		    robot.mousePress(InputEvent.BUTTON1_MASK);
		    robot.mouseRelease(InputEvent.BUTTON1_MASK);	
		    
		    Thread.sleep(SleepBtwClicks);
		    robot.keyPress(KeyEvent.VK_ENTER);
		    robot.keyRelease(KeyEvent.VK_ENTER);
			
		  //Enter Sale number	
		    
		    Thread.sleep(WaitForInstallPage); 
			MAX_X = 525;
			MAX_Y = 224;
		    robot.mouseMove(MAX_X, MAX_Y);
		    robot.mousePress(InputEvent.BUTTON1_MASK);
		    robot.mouseRelease(InputEvent.BUTTON1_MASK);
		    robot.mousePress(InputEvent.BUTTON1_MASK);
		    robot.mouseRelease(InputEvent.BUTTON1_MASK);
		    
		    Thread.sleep(SleepBtwClicks);
			robot.keyPress(KeyEvent.VK_8);
			robot.keyRelease(KeyEvent.VK_8);
			robot.keyPress(KeyEvent.VK_2);
			robot.keyRelease(KeyEvent.VK_2);
			robot.keyPress(KeyEvent.VK_7);
			robot.keyRelease(KeyEvent.VK_7);
			robot.keyPress(KeyEvent.VK_1);
			robot.keyRelease(KeyEvent.VK_1);
			robot.keyPress(KeyEvent.VK_4);
			robot.keyRelease(KeyEvent.VK_4);
			
		    Thread.sleep(SleepBtwClicks);
		    robot.keyPress(KeyEvent.VK_ENTER);
		    robot.keyRelease(KeyEvent.VK_ENTER);
		    
			//Enter Sale Type	    
		//	Thread.sleep(SleepBtwClicks);
			robot.keyPress(KeyEvent.VK_A);
			robot.keyRelease(KeyEvent.VK_A);
			robot.keyPress(KeyEvent.VK_A);
			robot.keyRelease(KeyEvent.VK_A);
			
			 //Press Enter to continue
		//    Thread.sleep(SleepBtwClicks);
		    robot.keyPress(KeyEvent.VK_ENTER);
		    robot.keyRelease(KeyEvent.VK_ENTER);
		      
			//Enter Sale Reason	    
		//	Thread.sleep(SleepBtwClicks);
			robot.keyPress(KeyEvent.VK_A);
			robot.keyRelease(KeyEvent.VK_A);
			robot.keyPress(KeyEvent.VK_X);
			robot.keyRelease(KeyEvent.VK_X);
			
		    //Press Enter to continue
	//	    Thread.sleep(SleepBtwClicks);
		    robot.keyPress(KeyEvent.VK_ENTER);
		    robot.keyRelease(KeyEvent.VK_ENTER);
		    
		    
		    //Enter Service Code - Page 1
		    Thread.sleep(SleepBtwClicks);
		    theString = ServiceCodeSet1;
		    StringSelection selection8 = new StringSelection(theString);
			Clipboard clipboard8 = Toolkit.getDefaultToolkit().getSystemClipboard();
			clipboard8.setContents(selection8, selection8);
		    //MOVE MOUSE AND PASTE
			MAX_X = 200;
			MAX_Y = 390;
		    robot.mouseMove(MAX_X, MAX_Y);
		    Thread.sleep(SleepBtwClicks);
		    robot.mousePress(InputEvent.BUTTON1_MASK);
		    robot.mouseRelease(InputEvent.BUTTON1_MASK);
			    
			Thread.sleep(SleepBtwClicks);
			robot.keyPress(KeyEvent.VK_CONTROL);
			robot.keyPress(KeyEvent.VK_V);
			robot.keyRelease(KeyEvent.VK_V);
			robot.keyRelease(KeyEvent.VK_CONTROL);
			Thread.sleep(SleepBtwClicks);
			
			//Enter Quantity Page-1 
		    theString = ServiceQTY1;
		    selection8 = new StringSelection(theString);
			clipboard8 = Toolkit.getDefaultToolkit().getSystemClipboard();
			clipboard8.setContents(selection8, selection8);
		    //MOVE MOUSE AND PASTE
			MAX_X = 400;
			MAX_Y = 390;
		    robot.mouseMove(MAX_X, MAX_Y);
		    Thread.sleep(SleepBtwClicks);
		    robot.mousePress(InputEvent.BUTTON1_MASK);
		    robot.mouseRelease(InputEvent.BUTTON1_MASK);
			    
			robot.keyPress(KeyEvent.VK_CONTROL);
			robot.keyPress(KeyEvent.VK_V);
			robot.keyRelease(KeyEvent.VK_V);
			robot.keyRelease(KeyEvent.VK_CONTROL);
			Thread.sleep(SleepBtwClicks);
			

		    robot.keyPress(KeyEvent.VK_ENTER);
		    robot.keyRelease(KeyEvent.VK_ENTER);
	

		    //Select packages
			
			MAX_X = 465;
			MAX_Y = 310;
		    robot.mouseMove(MAX_X, MAX_Y);
		    Thread.sleep(SleepBtwClicks);
		    robot.mousePress(InputEvent.BUTTON1_MASK);
		    robot.mouseRelease(InputEvent.BUTTON1_MASK);
		    robot.mousePress(InputEvent.BUTTON1_MASK);
		    robot.mouseRelease(InputEvent.BUTTON1_MASK);
		    
		    Thread.sleep(SleepBtwClicks);
			robot.keyPress(KeyEvent.VK_L);
			robot.keyRelease(KeyEvent.VK_L);
			robot.keyPress(KeyEvent.VK_N);
			robot.keyRelease(KeyEvent.VK_N);
			robot.keyPress(KeyEvent.VK_7);
			robot.keyRelease(KeyEvent.VK_7);
			robot.keyPress(KeyEvent.VK_1);
			robot.keyRelease(KeyEvent.VK_1);
			robot.keyPress(KeyEvent.VK_0);
			robot.keyRelease(KeyEvent.VK_0);
			
		    Thread.sleep(SleepBtwClicks);
		    robot.keyPress(KeyEvent.VK_ENTER);
		    robot.keyRelease(KeyEvent.VK_ENTER);
		    
		    
			 //Go to "41"
			Thread.sleep(SleepBtwClicks);
			MAX_X = 600;
			MAX_Y = 95;
		    robot.mouseMove(MAX_X, MAX_Y);
		    Thread.sleep(SleepBtwClicks);
		    robot.keyPress(KeyEvent.VK_4);
		    robot.keyRelease(KeyEvent.VK_4);
		    robot.keyPress(KeyEvent.VK_1);
		    robot.keyRelease(KeyEvent.VK_1);
		    
		    Thread.sleep(SleepBtwClicks);
		    robot.keyPress(KeyEvent.VK_ENTER);
		    robot.keyRelease(KeyEvent.VK_ENTER);
		    

		    
		    //click on "office only"
			Thread.sleep(SleepBtwClicks);
			MAX_X = 130;
			MAX_Y = 301;
		    robot.mouseMove(MAX_X, MAX_Y);
		    robot.mousePress(InputEvent.BUTTON1_MASK);
		    robot.mouseRelease(InputEvent.BUTTON1_MASK);
		    
		    
		    //Press Enter (s) to continue
		    Thread.sleep(SleepBtwClicks);
		    robot.keyPress(KeyEvent.VK_ENTER);
		    robot.keyRelease(KeyEvent.VK_ENTER);
		    robot.keyPress(KeyEvent.VK_ENTER);
		    robot.keyRelease(KeyEvent.VK_ENTER);
		    Thread.sleep(PageLoad);
			robot.keyPress(KeyEvent.VK_Y);
			robot.keyRelease(KeyEvent.VK_Y);
		    
		    
		    robot.keyPress(KeyEvent.VK_ENTER);
		    robot.keyRelease(KeyEvent.VK_ENTER);
		    robot.keyPress(KeyEvent.VK_ENTER);
		    robot.keyRelease(KeyEvent.VK_ENTER);
		    robot.keyPress(KeyEvent.VK_ENTER);
		    robot.keyRelease(KeyEvent.VK_ENTER);
		    
		    
			 //Go to "06"
		    Thread.sleep(PageLoad);
			MAX_X = 609;
			MAX_Y = 89;
		    robot.mouseMove(MAX_X, MAX_Y);
		    Thread.sleep(SleepBtwClicks);
		    robot.mousePress(InputEvent.BUTTON1_MASK);
		    robot.mouseRelease(InputEvent.BUTTON1_MASK);
		    robot.mousePress(InputEvent.BUTTON1_MASK);
		    robot.mouseRelease(InputEvent.BUTTON1_MASK);
		    robot.keyPress(KeyEvent.VK_0);
		    robot.keyRelease(KeyEvent.VK_0);
		    robot.keyPress(KeyEvent.VK_6);
		    robot.keyRelease(KeyEvent.VK_6);
			Thread.sleep(SleepBtwClicks);
		    robot.keyPress(KeyEvent.VK_ENTER);
		    robot.keyRelease(KeyEvent.VK_ENTER);
		    
		    Thread.sleep(2000);
			robot.keyPress(KeyEvent.VK_C);
			robot.keyRelease(KeyEvent.VK_C);
			robot.keyPress(KeyEvent.VK_U);
			robot.keyRelease(KeyEvent.VK_U);
			robot.keyPress(KeyEvent.VK_S);
			robot.keyRelease(KeyEvent.VK_S);
			robot.keyPress(KeyEvent.VK_T);
			robot.keyRelease(KeyEvent.VK_T);
			robot.keyPress(KeyEvent.VK_M);
			robot.keyRelease(KeyEvent.VK_M);
			robot.keyPress(KeyEvent.VK_O);
			robot.keyRelease(KeyEvent.VK_O);
			robot.keyPress(KeyEvent.VK_D);
			robot.keyRelease(KeyEvent.VK_D);
			Thread.sleep(SleepBtwClicks);
		    robot.keyPress(KeyEvent.VK_ENTER);
		    robot.keyRelease(KeyEvent.VK_ENTER);
		    Thread.sleep(SleepBtwClicks);
		    robot.keyPress(KeyEvent.VK_ENTER);
		    robot.keyRelease(KeyEvent.VK_ENTER);
		    
		    //click on completed
		    Thread.sleep(SleepBtwClicks);
		    MAX_X = 140;
			MAX_Y = 222;
		    robot.mouseMove(MAX_X, MAX_Y);
		    Thread.sleep(SleepBtwClicks);
		    robot.mousePress(InputEvent.BUTTON1_MASK);
		    robot.mouseRelease(InputEvent.BUTTON1_MASK);
		    robot.mousePress(InputEvent.BUTTON1_MASK);
		    robot.mouseRelease(InputEvent.BUTTON1_MASK);
		    
		    
		    //Enter Tech id
		    Thread.sleep(SleepBtwClicks);
		    MAX_X = 305;
			MAX_Y = 255;
		    robot.mouseMove(MAX_X, MAX_Y);
		    Thread.sleep(SleepBtwClicks);
		    robot.mousePress(InputEvent.BUTTON1_MASK);
		    robot.mouseRelease(InputEvent.BUTTON1_MASK);
		    robot.mousePress(InputEvent.BUTTON1_MASK);
		    robot.mouseRelease(InputEvent.BUTTON1_MASK);
		    Thread.sleep(SleepBtwClicks);
			robot.keyPress(KeyEvent.VK_8);
			robot.keyRelease(KeyEvent.VK_8);
			robot.keyPress(KeyEvent.VK_2);
			robot.keyRelease(KeyEvent.VK_2);
			robot.keyPress(KeyEvent.VK_7);
			robot.keyRelease(KeyEvent.VK_7);
			robot.keyPress(KeyEvent.VK_1);
			robot.keyRelease(KeyEvent.VK_1);
			robot.keyPress(KeyEvent.VK_4);
			robot.keyRelease(KeyEvent.VK_4);
		    Thread.sleep(SleepBtwClicks);
		    robot.keyPress(KeyEvent.VK_ENTER);
		    robot.keyRelease(KeyEvent.VK_ENTER);
		    
		    //Enter Solution code
		    Thread.sleep(SleepBtwClicks);
		    MAX_X = 279;
			MAX_Y = 409;
		    robot.mouseMove(MAX_X, MAX_Y);
		    Thread.sleep(SleepBtwClicks);
		    robot.mousePress(InputEvent.BUTTON1_MASK);
		    robot.mouseRelease(InputEvent.BUTTON1_MASK);
		    robot.mousePress(InputEvent.BUTTON1_MASK);
		    robot.mouseRelease(InputEvent.BUTTON1_MASK);
			robot.keyPress(KeyEvent.VK_C);
			robot.keyRelease(KeyEvent.VK_C);
			robot.keyPress(KeyEvent.VK_P);
			robot.keyRelease(KeyEvent.VK_P);
		    Thread.sleep(SleepBtwClicks);
		    robot.keyPress(KeyEvent.VK_ENTER);
		    robot.keyRelease(KeyEvent.VK_ENTER);
		    Thread.sleep(SleepBtwClicks);
		    robot.keyPress(KeyEvent.VK_ENTER);
		    robot.keyRelease(KeyEvent.VK_ENTER);
		    
		    
		    Thread.sleep(SleepBtwClicks);
			robot.keyPress(KeyEvent.VK_F3);
			robot.keyRelease(KeyEvent.VK_F3);

		}
	    }
	    
	    
	}
	
