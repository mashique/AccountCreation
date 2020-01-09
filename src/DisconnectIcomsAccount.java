
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

	class DisconnectIcomsAccount {
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
		FileInputStream f=new FileInputStream("D:\\Softwares\\Disconnect Icoms Account.xlsx");
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
		    
			selection3 = new StringSelection("61");
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
	    
	   
 
	      Thread.sleep(SmallWait);
	      robot.keyPress(KeyEvent.VK_ENTER);
	      robot.keyRelease(KeyEvent.VK_ENTER);
	      Thread.sleep(SmallWait);
	      robot.keyPress(KeyEvent.VK_ENTER);
	      robot.keyRelease(KeyEvent.VK_ENTER);
	      
	      
	      //Enter Reason code
	      Thread.sleep(WaitForInstallPage); 
			MAX_X = 237;
			MAX_Y = 236;
		    robot.mouseMove(MAX_X, MAX_Y);
		    robot.mousePress(InputEvent.BUTTON1_MASK);
		    robot.mouseRelease(InputEvent.BUTTON1_MASK);
		    robot.mousePress(InputEvent.BUTTON1_MASK);
		    robot.mouseRelease(InputEvent.BUTTON1_MASK);
		    
		    Thread.sleep(SleepBtwClicks);
			robot.keyPress(KeyEvent.VK_O);
			robot.keyRelease(KeyEvent.VK_O);
			robot.keyPress(KeyEvent.VK_Y);
			robot.keyRelease(KeyEvent.VK_Y);
			
		    Thread.sleep(SmallWait);
		    robot.keyPress(KeyEvent.VK_ENTER);
		    robot.keyRelease(KeyEvent.VK_ENTER);
			
			
			//Enter disconnect date
			
			Thread.sleep(WaitForInstallPage); 
			MAX_X = 241;
			MAX_Y = 258;
		    robot.mouseMove(MAX_X, MAX_Y);
		    robot.mousePress(InputEvent.BUTTON1_MASK);
		    robot.mouseRelease(InputEvent.BUTTON1_MASK);
		    robot.mousePress(InputEvent.BUTTON1_MASK);
		    robot.mouseRelease(InputEvent.BUTTON1_MASK);
		    
		    Thread.sleep(SmallWait);
			robot.keyPress(KeyEvent.VK_0);
			robot.keyRelease(KeyEvent.VK_0);
			robot.keyPress(KeyEvent.VK_6);
			robot.keyRelease(KeyEvent.VK_6);
			robot.keyPress(KeyEvent.VK_0);
			robot.keyRelease(KeyEvent.VK_0);
			robot.keyPress(KeyEvent.VK_3);
			robot.keyRelease(KeyEvent.VK_3);
			robot.keyPress(KeyEvent.VK_1);
			robot.keyRelease(KeyEvent.VK_1);
			robot.keyPress(KeyEvent.VK_9);
			robot.keyRelease(KeyEvent.VK_9);
			
		    Thread.sleep(SmallWait);
		    robot.keyPress(KeyEvent.VK_ENTER);
		    robot.keyRelease(KeyEvent.VK_ENTER);
		    
		    
		    //select office only
			
			Thread.sleep(SmallWait); 
			MAX_X = 127;
			MAX_Y = 302;
		    robot.mouseMove(MAX_X, MAX_Y);
		    robot.mousePress(InputEvent.BUTTON1_MASK);
		    robot.mouseRelease(InputEvent.BUTTON1_MASK);

		    
		    Thread.sleep(SmallWait);
		    robot.keyPress(KeyEvent.VK_ENTER);
		    robot.keyRelease(KeyEvent.VK_ENTER);
		    Thread.sleep(SmallWait);
		    robot.keyPress(KeyEvent.VK_ENTER);
		    robot.keyRelease(KeyEvent.VK_ENTER);
		    
		    Thread.sleep(SleepBtwClicks);
			robot.keyPress(KeyEvent.VK_Y);
			robot.keyRelease(KeyEvent.VK_Y);
			
		    Thread.sleep(SmallWait);
		    robot.keyPress(KeyEvent.VK_ENTER);
		    robot.keyRelease(KeyEvent.VK_ENTER);		
			
		    Thread.sleep(SmallWait);
		    robot.keyPress(KeyEvent.VK_ENTER);
		    robot.keyRelease(KeyEvent.VK_ENTER);
		    
	      
		  Thread.sleep(2000);
		  
		  //Enter Bill to address
		  Thread.sleep(SleepBtwClicks);
		   	MAX_X = 531;
			MAX_Y = 286;
		    robot.mouseMove(MAX_X, MAX_Y);
		  
		    robot.mousePress(InputEvent.BUTTON1_MASK);
		    robot.mouseRelease(InputEvent.BUTTON1_MASK);
		    robot.mousePress(InputEvent.BUTTON1_MASK);
		    robot.mouseRelease(InputEvent.BUTTON1_MASK);
		    
			selection3 = new StringSelection("6380");
			clipboard3 = Toolkit.getDefaultToolkit().getSystemClipboard();
			clipboard3.setContents(selection3, selection3);
		    
			Thread.sleep(SleepBtwClicks);
		      robot.keyPress(KeyEvent.VK_CONTROL);
		      robot.keyPress(KeyEvent.VK_V);
		      robot.keyRelease(KeyEvent.VK_V);
		      robot.keyRelease(KeyEvent.VK_CONTROL);

	//---------------------------------------------------------
		      
		      Thread.sleep(SleepBtwClicks);
		      MAX_X = 602;
				MAX_Y = 279;
			    robot.mouseMove(MAX_X, MAX_Y);
			  
			    robot.mousePress(InputEvent.BUTTON1_MASK);
			    robot.mouseRelease(InputEvent.BUTTON1_MASK);
			    robot.mousePress(InputEvent.BUTTON1_MASK);
			    robot.mouseRelease(InputEvent.BUTTON1_MASK);
			    
				selection3 = new StringSelection("S");
				clipboard3 = Toolkit.getDefaultToolkit().getSystemClipboard();
				clipboard3.setContents(selection3, selection3);
			    
				Thread.sleep(SleepBtwClicks);
			      robot.keyPress(KeyEvent.VK_CONTROL);
			      robot.keyPress(KeyEvent.VK_V);
			      robot.keyRelease(KeyEvent.VK_V);
			      robot.keyRelease(KeyEvent.VK_CONTROL);
			      
	//--------------------------------------------------------------------------		      
			      Thread.sleep(SleepBtwClicks);
			      MAX_X = 662;
					MAX_Y = 274;
				    robot.mouseMove(MAX_X, MAX_Y);
				  
				    robot.mousePress(InputEvent.BUTTON1_MASK);
				    robot.mouseRelease(InputEvent.BUTTON1_MASK);
				    robot.mousePress(InputEvent.BUTTON1_MASK);
				    robot.mouseRelease(InputEvent.BUTTON1_MASK);
				    
					selection3 = new StringSelection("FIDDLERS GREEN CIR FL 9");
					clipboard3 = Toolkit.getDefaultToolkit().getSystemClipboard();
					clipboard3.setContents(selection3, selection3);
				    
					Thread.sleep(SleepBtwClicks);
				      robot.keyPress(KeyEvent.VK_CONTROL);
				      robot.keyPress(KeyEvent.VK_V);
				      robot.keyRelease(KeyEvent.VK_V);
				      robot.keyRelease(KeyEvent.VK_CONTROL);
		      
	//---------------------------------------------------------------------------------------------			      
		      
				      Thread.sleep(SleepBtwClicks);
				      MAX_X = 507;
						MAX_Y = 391;
					    robot.mouseMove(MAX_X, MAX_Y);
					  
					    robot.mousePress(InputEvent.BUTTON1_MASK);
					    robot.mouseRelease(InputEvent.BUTTON1_MASK);
					    robot.mousePress(InputEvent.BUTTON1_MASK);
					    robot.mouseRelease(InputEvent.BUTTON1_MASK);
					    
						selection3 = new StringSelection("GREENWOOD VILLAGE");
						clipboard3 = Toolkit.getDefaultToolkit().getSystemClipboard();
						clipboard3.setContents(selection3, selection3);
					    
						Thread.sleep(SleepBtwClicks);
					      robot.keyPress(KeyEvent.VK_CONTROL);
					      robot.keyPress(KeyEvent.VK_V);
					      robot.keyRelease(KeyEvent.VK_V);
					      robot.keyRelease(KeyEvent.VK_CONTROL);
		    //------------------------------------------------------------
					      
					      Thread.sleep(SleepBtwClicks);
					      MAX_X = 657;
							MAX_Y = 390;
						    robot.mouseMove(MAX_X, MAX_Y);
						  
						    robot.mousePress(InputEvent.BUTTON1_MASK);
						    robot.mouseRelease(InputEvent.BUTTON1_MASK);
						    robot.mousePress(InputEvent.BUTTON1_MASK);
						    robot.mouseRelease(InputEvent.BUTTON1_MASK);
						    
							selection3 = new StringSelection("co");
							clipboard3 = Toolkit.getDefaultToolkit().getSystemClipboard();
							clipboard3.setContents(selection3, selection3);
						    
							Thread.sleep(SleepBtwClicks);
						      robot.keyPress(KeyEvent.VK_CONTROL);
						      robot.keyPress(KeyEvent.VK_V);
						      robot.keyRelease(KeyEvent.VK_V);
						      robot.keyRelease(KeyEvent.VK_CONTROL);
						      
						      
			//----------------------------------------------------------------	
	
						      
						      Thread.sleep(SleepBtwClicks);
						      MAX_X = 704;
								MAX_Y = 383;
							    robot.mouseMove(MAX_X, MAX_Y);
							  
							    robot.mousePress(InputEvent.BUTTON1_MASK);
							    robot.mouseRelease(InputEvent.BUTTON1_MASK);
							    robot.mousePress(InputEvent.BUTTON1_MASK);
							    robot.mouseRelease(InputEvent.BUTTON1_MASK);
							    
								selection3 = new StringSelection("80111");
								clipboard3 = Toolkit.getDefaultToolkit().getSystemClipboard();
								clipboard3.setContents(selection3, selection3);
							    
								Thread.sleep(SleepBtwClicks);
							      robot.keyPress(KeyEvent.VK_CONTROL);
							      robot.keyPress(KeyEvent.VK_V);
							      robot.keyRelease(KeyEvent.VK_V);
							      robot.keyRelease(KeyEvent.VK_CONTROL);
			//----------------------------------------------------------------------
							      
							        Thread.sleep(SmallWait); 
									MAX_X = 869;
									MAX_Y = 239;
								    robot.mouseMove(MAX_X, MAX_Y);
								    robot.mousePress(InputEvent.BUTTON1_MASK);
								    robot.mouseRelease(InputEvent.BUTTON1_MASK);
								    
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
			robot.keyPress(KeyEvent.VK_F3);
			robot.keyRelease(KeyEvent.VK_F3);

		}
	    }
	    
	    
	}
	
