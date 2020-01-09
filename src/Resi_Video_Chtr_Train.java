
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.event.InputEvent;
import java.awt.event.KeyEvent;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.StringSelection;
import java.awt.*;
import java.awt.event.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



	public class Resi_Video_Chtr_Train {
    
	    public static final int SleepBtwClicks = 1500;
	    public static final int NewPageLoad = 5000;
	    public static int MAX_Y = 440;
	    public static int MAX_X = 110;
	    public static int StartExecution = 2000;
	    public static String FirstName;
	    public static String LastName;
	    public static String LocationID;
	    public static String AccountNumber;
	    public static String OrderID;
	    
	    public static void main(String... args) throws Exception {
	    Robot robot = new Robot();

		FileInputStream f=new FileInputStream("D:\\Softwares\\Charter Training Account Creation.xlsx");
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
			LastName=cell2.getStringCellValue();
			Cell cell3=s.getRow(row).getCell(2);
			FirstName=cell3.getStringCellValue();
			
			
					
		
		Thread.sleep(StartExecution);
//		Toolkit toolkit = Toolkit.getDefaultToolkit();
//		Clipboard clipboard = toolkit.getSystemClipboard();

		StringSelection selection = new StringSelection(LocationID);
		Clipboard clipboard1 = Toolkit.getDefaultToolkit().getSystemClipboard();
		clipboard1.setContents(selection, selection);
	    Thread.sleep(SleepBtwClicks);
	    
		Thread.sleep(NewPageLoad);
		Thread.sleep(SleepBtwClicks);
		robot.keyPress(KeyEvent.VK_ALT);
		robot.keyPress(KeyEvent.VK_C);
		robot.keyRelease(KeyEvent.VK_C);
		robot.keyRelease(KeyEvent.VK_ALT);
		
		Thread.sleep(NewPageLoad);
		Thread.sleep(SleepBtwClicks);
		robot.keyPress(KeyEvent.VK_ALT);
		robot.keyPress(KeyEvent.VK_C);
		robot.keyRelease(KeyEvent.VK_C);
		robot.keyRelease(KeyEvent.VK_ALT);
		Thread.sleep(SleepBtwClicks);
		Thread.sleep(SleepBtwClicks);
		
	  /*  //Move Cursor to hit Clear
		MAX_X = 1320;
		MAX_Y = 676;
	    robot.mouseMove(MAX_X, MAX_Y);
	    Thread.sleep(SleepBtwClicks);
	    robot.mousePress(InputEvent.BUTTON1_MASK);
	    robot.mouseRelease(InputEvent.BUTTON1_MASK);
		*/
		
	    //Move Cursor to enter LocationID
		MAX_X = 753;
		MAX_Y = 406	;
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
	      Thread.sleep(SleepBtwClicks);
		selection = new StringSelection(LastName);
		clipboard1 = Toolkit.getDefaultToolkit().getSystemClipboard();
		clipboard1.setContents(selection, selection);
	    Thread.sleep(SleepBtwClicks);
			robot.keyPress(KeyEvent.VK_CONTROL);
			robot.keyPress(KeyEvent.VK_V);
			robot.keyRelease(KeyEvent.VK_V);
			robot.keyRelease(KeyEvent.VK_CONTROL);
//			Thread.sleep(SleepBtwClicks);
//			robot.keyPress(KeyEvent.VK_TAB);
//		    robot.keyRelease(KeyEvent.VK_TAB);
		
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
		
				
	    //Hit Enter - Twice
		Thread.sleep(SleepBtwClicks);
		robot.keyPress(KeyEvent.VK_ENTER);
		robot.keyRelease(KeyEvent.VK_ENTER);
		Thread.sleep(SleepBtwClicks);
		robot.keyPress(KeyEvent.VK_ENTER);
		robot.keyRelease(KeyEvent.VK_ENTER);	
		
	    //Move Cursor to select "Standard Rate" tab
		Thread.sleep(NewPageLoad);
		MAX_X = 73;
		MAX_Y = 176;
	    robot.mouseMove(MAX_X, MAX_Y);
	    Thread.sleep(SleepBtwClicks);
	    robot.mousePress(InputEvent.BUTTON1_MASK);
	    robot.mouseRelease(InputEvent.BUTTON1_MASK);
			   
	    //Move Cursor to select required package
		Thread.sleep(SleepBtwClicks);
		MAX_X = 321;
		MAX_Y = 408;
	    robot.mouseMove(MAX_X, MAX_Y);
	    Thread.sleep(SleepBtwClicks);
	    robot.mousePress(InputEvent.BUTTON1_MASK);
	    robot.mouseRelease(InputEvent.BUTTON1_MASK);
	    
	    //Click on ADD package
		Thread.sleep(SleepBtwClicks);
		MAX_X = 1272;
		MAX_Y = 454;
	    robot.mouseMove(MAX_X, MAX_Y);
	    Thread.sleep(SleepBtwClicks);
	    robot.mousePress(InputEvent.BUTTON1_MASK);
	    robot.mouseRelease(InputEvent.BUTTON1_MASK);
	    
	    //Click on Next Decision
		Thread.sleep(SleepBtwClicks);
		MAX_X = 1292;
		MAX_Y = 109;
	    robot.mouseMove(MAX_X, MAX_Y);
	    Thread.sleep(SleepBtwClicks);
	    robot.mousePress(InputEvent.BUTTON1_MASK);
	    robot.mouseRelease(InputEvent.BUTTON1_MASK);
	    Thread.sleep(500);
	    
	    
	    //Scroll
	    Thread.sleep(SleepBtwClicks);
		MAX_X = 1348;
		MAX_Y = 350;
	    robot.mouseMove(MAX_X, MAX_Y);
	    Thread.sleep(SleepBtwClicks);
	    robot.mousePress(InputEvent.BUTTON1_MASK);
  	    robot.mouseRelease(InputEvent.BUTTON1_MASK);
  	    robot.mousePress(InputEvent.BUTTON1_MASK);
  	    robot.mouseRelease(InputEvent.BUTTON1_MASK);
  	    
  	    
  	    //Select HD Box
  	    Thread.sleep(SleepBtwClicks);
		MAX_X = 1074;
		MAX_Y = 347;
	    robot.mouseMove(MAX_X, MAX_Y);
	    Thread.sleep(SleepBtwClicks);
	    robot.mousePress(InputEvent.BUTTON1_MASK);
	    robot.mouseRelease(InputEvent.BUTTON1_MASK);
	    
	    
	  //Click on Next Decision
	  		Thread.sleep(SleepBtwClicks);
	  		MAX_X = 1292;
	  		MAX_Y = 109;
	  	    robot.mouseMove(MAX_X, MAX_Y);
	  	    Thread.sleep(SleepBtwClicks);
	  	    for (int i =0; i<5;i++)
	  	    {
	  	    robot.mousePress(InputEvent.BUTTON1_MASK);
	  	    robot.mouseRelease(InputEvent.BUTTON1_MASK);
	  	    Thread.sleep(500);
	  	    }
	    
	    
	    //Click on SAVE
		Thread.sleep(SleepBtwClicks);
		MAX_X = 1070;
		MAX_Y = 706;
	    robot.mouseMove(MAX_X, MAX_Y);
	    Thread.sleep(SleepBtwClicks);
	    robot.mousePress(InputEvent.BUTTON1_MASK);
	    robot.mouseRelease(InputEvent.BUTTON1_MASK);
	    Thread.sleep(SleepBtwClicks);
	    
	    //Click on Next
		MAX_X = 1245;
		MAX_Y = 707;
	    robot.mouseMove(MAX_X, MAX_Y);
	    Thread.sleep(SleepBtwClicks);
	    robot.mousePress(InputEvent.BUTTON1_MASK);
	    robot.mouseRelease(InputEvent.BUTTON1_MASK);
	    Thread.sleep(SleepBtwClicks);
	    robot.mousePress(InputEvent.BUTTON1_MASK);
	    robot.mouseRelease(InputEvent.BUTTON1_MASK);
	    
	    //Click on Finish
		Thread.sleep(SleepBtwClicks);
		MAX_X = 1245;
		MAX_Y = 707;
	    robot.mouseMove(MAX_X, MAX_Y);
	    Thread.sleep(SleepBtwClicks);
	    robot.mousePress(InputEvent.BUTTON1_MASK);
	    robot.mouseRelease(InputEvent.BUTTON1_MASK);
	    Thread.sleep(SleepBtwClicks);
	    
	    
	    //Next page Alt+N
	    Thread.sleep(SleepBtwClicks);
		robot.keyPress(KeyEvent.VK_ALT);
		robot.keyPress(KeyEvent.VK_N);
		robot.keyRelease(KeyEvent.VK_N);
		robot.keyRelease(KeyEvent.VK_ALT);
		
	    //Enter Tax Type
		Thread.sleep(NewPageLoad);
		MAX_X = 763;
		MAX_Y = 282;
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
		robot.keyPress(KeyEvent.VK_8);
		robot.keyRelease(KeyEvent.VK_8);

		
	    //Next page Alt+N
	    Thread.sleep(SleepBtwClicks);
		robot.keyPress(KeyEvent.VK_ALT);
		robot.keyPress(KeyEvent.VK_N);
		robot.keyRelease(KeyEvent.VK_N);
		robot.keyRelease(KeyEvent.VK_ALT);
		
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
		robot.keyPress(KeyEvent.VK_0);
		robot.keyRelease(KeyEvent.VK_0);
		robot.keyPress(KeyEvent.VK_0);
		robot.keyRelease(KeyEvent.VK_0);
		robot.keyPress(KeyEvent.VK_0);
		robot.keyRelease(KeyEvent.VK_0);
		robot.keyPress(KeyEvent.VK_0);
		robot.keyRelease(KeyEvent.VK_0);
		Thread.sleep(SleepBtwClicks);
		robot.keyPress(KeyEvent.VK_TAB);
		robot.keyRelease(KeyEvent.VK_TAB);
		Thread.sleep(SleepBtwClicks);
		robot.keyPress(KeyEvent.VK_X);
		robot.keyRelease(KeyEvent.VK_X);
		robot.keyPress(KeyEvent.VK_X);
		robot.keyRelease(KeyEvent.VK_X);
		
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
		robot.keyPress(KeyEvent.VK_N);
		robot.keyRelease(KeyEvent.VK_N);
		robot.keyRelease(KeyEvent.VK_ALT);
		
		
		
		//Enter to Finish
		Thread.sleep(NewPageLoad);
		robot.keyPress(KeyEvent.VK_ENTER);
		robot.keyRelease(KeyEvent.VK_ENTER);
		Thread.sleep(NewPageLoad);
		robot.keyPress(KeyEvent.VK_ENTER);
		robot.keyRelease(KeyEvent.VK_ENTER);
		
		//Enter Alt+F4 to close contact window
		Thread.sleep(SleepBtwClicks);
		Thread.sleep(NewPageLoad);
		robot.keyPress(KeyEvent.VK_ALT);
		robot.keyPress(KeyEvent.VK_F4);
		robot.keyRelease(KeyEvent.VK_F4);
		robot.keyRelease(KeyEvent.VK_ALT);
				
		
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
		FileOutputStream fo=new FileOutputStream("D:\\Softwares\\Charter Training Account Creation.xlsx");
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
		Cell cell4=s.getRow(row).createCell(s.getRow(row).getPhysicalNumberOfCells());
		cell4.setCellType(cell4.CELL_TYPE_STRING);
		cell4.setCellValue(OrderID);
		Thread.sleep(SleepBtwClicks);
		wb.write(fo);
		
	    
	
		}
			catch(Exception e){
				
				Cell cell=s.getRow(row).createCell(s.getRow(row).getPhysicalNumberOfCells());
				cell.setCellType(cell.CELL_TYPE_STRING);
				cell.setCellValue(e.toString());
				FileOutputStream fo=new FileOutputStream("D:\\Softwares\\Charter Training Account Creation.xlsx");
				wb.write(fo);
					
						}
		}
	    }
	    
	   
	}
	
