
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



	public class DeviceIcoms {
    
	    public static final int SleepBtwClicks = 1500;
	    public static final int NewPageLoad = 5000;
	    public static int MAX_Y = 440;
	    public static int MAX_X = 110;
	    public static int StartExecution = 2000;
	    public static String Serial;
	    public static String Mac1;
	    public static String Mac2;

	    
	    public static void main(String... args) throws Exception {
	    Robot robot = new Robot();

		FileInputStream f=new FileInputStream("D:\\Softwares\\Device_Entry.xlsx");
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
			Serial=cell1.getStringCellValue();
			Cell cell2=s.getRow(row).getCell(1);
			Mac1=cell2.getStringCellValue();
			Cell cell3=s.getRow(row).getCell(2);
			Mac2=cell3.getStringCellValue();
			
					
		    //Enter Serial number 
			Thread.sleep(SleepBtwClicks);
			MAX_X = 211;
			MAX_Y = 85;
		    robot.mouseMove(MAX_X, MAX_Y);
		  
		    Thread.sleep(SleepBtwClicks);
		    robot.mousePress(InputEvent.BUTTON1_MASK);
		    robot.mouseRelease(InputEvent.BUTTON1_MASK);
		    robot.mousePress(InputEvent.BUTTON1_MASK);
		    robot.mouseRelease(InputEvent.BUTTON1_MASK);
		    
			StringSelection selection3 = new StringSelection(Serial);
			Clipboard clipboard3 = Toolkit.getDefaultToolkit().getSystemClipboard();
			clipboard3.setContents(selection3, selection3);
		    
			Thread.sleep(SleepBtwClicks);
		    robot.keyPress(KeyEvent.VK_CONTROL);
		    robot.keyPress(KeyEvent.VK_V);
		    robot.keyRelease(KeyEvent.VK_V);
		    robot.keyRelease(KeyEvent.VK_CONTROL);
		    Thread.sleep(SleepBtwClicks);
		    
		    
		    //Enter  MAC1 
			Thread.sleep(SleepBtwClicks);
			MAX_X = 211;
			MAX_Y = 85;
		    robot.mouseMove(MAX_X, MAX_Y);
		  
		    Thread.sleep(SleepBtwClicks);
		    robot.mousePress(InputEvent.BUTTON1_MASK);
		    robot.mouseRelease(InputEvent.BUTTON1_MASK);
		    robot.mousePress(InputEvent.BUTTON1_MASK);
		    robot.mouseRelease(InputEvent.BUTTON1_MASK);
		    
			selection3 = new StringSelection(Mac1);
			clipboard3 = Toolkit.getDefaultToolkit().getSystemClipboard();
			clipboard3.setContents(selection3, selection3);
		    
			Thread.sleep(SleepBtwClicks);
		    robot.keyPress(KeyEvent.VK_CONTROL);
		    robot.keyPress(KeyEvent.VK_V);
		    robot.keyRelease(KeyEvent.VK_V);
		    robot.keyRelease(KeyEvent.VK_CONTROL);
		    Thread.sleep(SleepBtwClicks);

		    //Enter  MAC2 
			Thread.sleep(SleepBtwClicks);
			MAX_X = 211;
			MAX_Y = 85;
		    robot.mouseMove(MAX_X, MAX_Y);
		  
		    Thread.sleep(SleepBtwClicks);
		    robot.mousePress(InputEvent.BUTTON1_MASK);
		    robot.mouseRelease(InputEvent.BUTTON1_MASK);
		    robot.mousePress(InputEvent.BUTTON1_MASK);
		    robot.mouseRelease(InputEvent.BUTTON1_MASK);
		    
			selection3 = new StringSelection(Mac2);
			clipboard3 = Toolkit.getDefaultToolkit().getSystemClipboard();
			clipboard3.setContents(selection3, selection3);
		    
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