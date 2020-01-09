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



	public class LocationCreation {
    
	    public static final int SleepBtwClicks = 1500;
	    public static final int NewPageLoad = 5000;
	    public static int MAX_Y = 440;
	    public static int MAX_X = 110;
	    public static int StartExecution = 2000;
	    public static String Address;
	    public static String AccountNumber;
	    public static String OrderID;
	    
	    public static void main(String... args) throws Exception {
	    Robot robot = new Robot();

		FileInputStream f=new FileInputStream("D:\\Softwares\\LocationCreation.xlsx");
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
			Address =cell1.getStringCellValue();
			
		
			
		
		Thread.sleep(StartExecution);
//		Toolkit toolkit = Toolkit.getDefaultToolkit();
//		Clipboard clipboard = toolkit.getSystemClipboard();

		StringSelection selection = new StringSelection(Address);
		Clipboard clipboard1 = Toolkit.getDefaultToolkit().getSystemClipboard();
		clipboard1.setContents(selection, selection);
	    Thread.sleep(SleepBtwClicks);
	    
		Thread.sleep(NewPageLoad);
		
		
	    //Move Cursor to enter Add2
		MAX_X = 189;
		MAX_Y = 213;
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
		Thread.sleep(SleepBtwClicks);
		robot.keyPress(KeyEvent.VK_ENTER);
		robot.keyRelease(KeyEvent.VK_ENTER);
//		Thread.sleep(SleepBtwClicks);
//		robot.keyPress(KeyEvent.VK_ENTER);
//		robot.keyRelease(KeyEvent.VK_ENTER);


		
	
	    //Copy Location and write in spreadsheet
		Thread.sleep(NewPageLoad);
	
		MAX_X = 117;
		MAX_Y = 100;
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
	       
		//Copy the Location from Clip board to String
		Thread.sleep(SleepBtwClicks);
		Toolkit toolkit = Toolkit.getDefaultToolkit();
		Clipboard clipboard = toolkit.getSystemClipboard();
		String AccountNumber = (String) clipboard.getData(DataFlavor.stringFlavor);
		 
		
		//Save the AccountNumber in spreadsheet
		Cell cell=s.getRow(row).createCell(s.getRow(row).getPhysicalNumberOfCells());
		cell.setCellType(cell.CELL_TYPE_STRING);
		cell.setCellValue(AccountNumber);
		Thread.sleep(SleepBtwClicks);
		FileOutputStream fo=new FileOutputStream("D:\\Softwares\\LocationCreation.xlsx");
		wb.write(fo);
		
	
		}
			catch(Exception e){
				
				Cell cell=s.getRow(row).createCell(s.getRow(row).getPhysicalNumberOfCells());
				cell.setCellType(cell.CELL_TYPE_STRING);
				cell.setCellValue(e.toString());
				FileOutputStream fo=new FileOutputStream("D:\\Softwares\\LocationCreation.xlsx");
				wb.write(fo);
					
						}
		}
	    }
	    
	   
	}
	