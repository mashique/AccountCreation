


import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.awt.AWTException;
import java.awt.Rectangle;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.event.InputEvent;
import java.awt.event.KeyEvent;
import java.util.Random;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.StringSelection;
import java.awt.*;
import java.awt.event.*;
import java.awt.image.BufferedImage;

import javax.imageio.ImageIO;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



	public class Update_Bill_To_Addr_LTWC_PROD extends Dialog implements ActionListener {
		private Button ok,can;
	    public boolean isOk = false;

	    /*
	     * @param frame   parent frame 
	     * @param msg     message to be displayed
	     * @param okcan   true : ok cancel buttons, false : ok button only 
	     */
	    Update_Bill_To_Addr_LTWC_PROD(Frame frame, String msg, boolean okcan){
	        super(frame, "Message", true);
	        setLayout(new BorderLayout());
	        add("Center",new Label(msg));
	        addOKCancelPanel(okcan);
	        createFrame();
	        pack();
	        setVisible(true);
	    }
	    
	    Update_Bill_To_Addr_LTWC_PROD(Frame frame, String msg){
	        this(frame, msg, false);
	    }
	    
	    void addOKCancelPanel( boolean okcan ) {
	        Panel p = new Panel();
	        p.setLayout(new FlowLayout());
	        createOKButton( p );
	        if (okcan == true)
	            createCancelButton( p );
	        add("South",p);
	    }

	    void createOKButton(Panel p) {
	        p.add(ok = new Button("OK"));
	        ok.addActionListener(this); 
	    }

	    void createCancelButton(Panel p) {
	        p.add(can = new Button("Cancel"));
	        can.addActionListener(this);
	    }

	    void createFrame() {
	        Dimension d = getToolkit().getScreenSize();
	        setLocation(d.width/3,d.height/3);
	    }

	    public void actionPerformed(ActionEvent ae){
	        if(ae.getSource() == ok) {
	            isOk = true;
	            setVisible(false);
	        }
	        else if(ae.getSource() == can) {
	            setVisible(false);
	        }
	    }
		
		
	    public static final int FIVE_SECONDS = 2000;
	    public static final int SleepBtwClicks = 2000;
	    public static final int NewPageLoad = 5000;
	    public static boolean ImageMatched;
	  //public static final int MAX_Y = 5;
	  //public static final int MAX_X = 5;
	    public static int MAX_Y = 440;
	    public static int MAX_X = 110;
	    public static int StartExecution = 2000;
	    public static int count = 0;
	    public static String theString = "123445";
	    public static String Location;
	    public static String Address_line1 = "6380 S Fiddler’s GreenCir";
	    public static String Address_line2 = "Greenwood Village";
	    public static String Address_line3 = "CO";
	    public static String Address_line4 = "80111";
	    public static String Memo = "Tamtool test account.\n For any questions related to this test\n account please contact DL-TDM";
	    
	    public static void main(String... args) throws Exception {
		Robot robot = new Robot();
		//Random random = new Random();
		//MAX_X  = OFFSET; 
		//MAX_Y = OFFSET;
		FileInputStream f=new FileInputStream("C:\\Driver\\UpdateProdCSGAccounts.xlsx");
		XSSFWorkbook wb = new XSSFWorkbook(f);
		XSSFSheet s= wb.getSheet("Sheet1");
		int totalRow=s.getPhysicalNumberOfRows();
		
		System.out.println("number of rows in spreadsheet"+totalRow);
		for(int row=1; row<totalRow; row++){
			Cell cell1=s.getRow(row).getCell(0);
			//Cell cell2=s.getRow(row).getCell(1);
			//Cell cell3=s.getRow(row).getCell(2);
			//Cell cell4=s.getRow(row).getCell(3);
			Location=cell1.getStringCellValue();
			//VideoMAC=cell2.getStringCellValue();
			//eSTB=cell3.getStringCellValue();
			//eCMMAC=cell4.getStringCellValue();
		
		//theString = Location;
		
		Thread.sleep(StartExecution);
//		Toolkit toolkit = Toolkit.getDefaultToolkit();
//		Clipboard clipboard = toolkit.getSystemClipboard();

		StringSelection selection = new StringSelection(Location);
		Clipboard clipboard1 = Toolkit.getDefaultToolkit().getSystemClipboard();
		clipboard1.setContents(selection, selection);
		System.out.println("inside while loop "+row);
		System.out.println("Location ID"+Location);
		
	    //Hit Clear
		MAX_X = 1284;
		MAX_Y = 674;
		//MAX_Y = 710;
	    robot.mouseMove(MAX_X, MAX_Y);
	    Thread.sleep(1000);
	    robot.mousePress(InputEvent.BUTTON1_MASK);
	    robot.mouseRelease(InputEvent.BUTTON1_MASK);
	    robot.mousePress(InputEvent.BUTTON1_MASK);
	    robot.mouseRelease(InputEvent.BUTTON1_MASK);
		
	  //*Validate to be on correct page to proceed - First page
	  		ImageMatched = false;
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
	  	           Rectangle rectArea = new Rectangle(1080, 668, 88, 27);
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

	  	       System.out.println("width and hights"+"-"+width1+"-"+width2+"-"+height1+"-"+height2);
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
	  	           if (difference <= 8000){
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
	   
	    //Enter Account
	    //location for resi account number
	    MAX_X = 448;
	    MAX_Y = 269;
	    //location for commercial account number
	    //MAX_X = 469;
	    //MAX_Y = 278;

	    
	    //Enter Location
	    //location for residential
		//MAX_X = 450;
		//MAX_Y = 495;
		//location for commercial 
		//MAX_X = 455;
		//MAX_Y = 468;
	    robot.mouseMove(MAX_X, MAX_Y);
	    Thread.sleep(2000);
	    robot.mousePress(InputEvent.BUTTON1_MASK);
	    robot.mouseRelease(InputEvent.BUTTON1_MASK);
	    robot.mousePress(InputEvent.BUTTON1_MASK);
	    robot.mouseRelease(InputEvent.BUTTON1_MASK);
	    
	    
	    //PressPaste2();
	    Thread.sleep(3000);
		 robot.keyPress(KeyEvent.VK_CONTROL);
		 robot.keyPress(KeyEvent.VK_V);
		 robot.keyRelease(KeyEvent.VK_V);
		 robot.keyRelease(KeyEvent.VK_CONTROL);
	    
	    //Hit Search
		MAX_X = 1116;
		MAX_Y = 674;
		//MAX_Y = 710;
	    robot.mouseMove(MAX_X, MAX_Y);
	    Thread.sleep(2000);
	    robot.mousePress(InputEvent.BUTTON1_MASK);
	    robot.mouseRelease(InputEvent.BUTTON1_MASK);
	    robot.mousePress(InputEvent.BUTTON1_MASK);
	    robot.mouseRelease(InputEvent.BUTTON1_MASK);
	    
	   
	    
	
	    //Click on "Account Information"
	    Thread.sleep(5000);
	    MAX_X = 154;
		MAX_Y = 55;
	    robot.mouseMove(MAX_X, MAX_Y);
	    	   
	    robot.mousePress(InputEvent.BUTTON1_MASK);
	    robot.mouseRelease(InputEvent.BUTTON1_MASK);
	    

		 //Click on "VIP"
		 Thread.sleep(2000);
		 MAX_X = 645;
		 MAX_Y = 246;
		 robot.mouseMove(MAX_X, MAX_Y);
		 	   
		 robot.mousePress(InputEvent.BUTTON1_MASK);
		 robot.mouseRelease(InputEvent.BUTTON1_MASK);
		 
		 
		 //Press 8 for entering VIP code
		 Thread.sleep(3000);
		 robot.keyPress(KeyEvent.VK_J);
		 robot.keyRelease(KeyEvent.VK_J);
		
	    //Click on "Profile"
	    Thread.sleep(3000);
	    MAX_X = 103;
		MAX_Y = 85;
	    robot.mouseMove(MAX_X, MAX_Y);
	    	   
	    robot.mousePress(InputEvent.BUTTON1_MASK);
	    robot.mouseRelease(InputEvent.BUTTON1_MASK);
	    
	    //Select Address line1
	    Thread.sleep(2000);
	    MAX_X = 187;
		MAX_Y = 562;
	    robot.mouseMove(MAX_X, MAX_Y);
	    	   
	    robot.mousePress(InputEvent.BUTTON1_MASK);
	    
	    MAX_X = 412;
		MAX_Y = 558;
	    robot.mouseMove(MAX_X, MAX_Y);
	    robot.mouseRelease(InputEvent.BUTTON1_MASK);
	    
		StringSelection selection2 = new StringSelection(Address_line1);
		Clipboard clipboard = Toolkit.getDefaultToolkit().getSystemClipboard();
		clipboard.setContents(selection2, selection2);
		   
	      robot.keyPress(KeyEvent.VK_CONTROL);
	      robot.keyPress(KeyEvent.VK_V);
	      robot.keyRelease(KeyEvent.VK_V);
	      robot.keyRelease(KeyEvent.VK_CONTROL);
	      
	    //blank out next address line  
		MAX_X = 400;
		MAX_Y = 569;
		robot.mouseMove(MAX_X, MAX_Y);
        robot.keyPress(KeyEvent.VK_SPACE);
	    robot.keyRelease(KeyEvent.VK_SPACE);
	    
	  //Select Address line City
	    Thread.sleep(2000);
	    MAX_X = 186;
		MAX_Y = 611;
	    robot.mouseMove(MAX_X, MAX_Y);
	    	   
	    robot.mousePress(InputEvent.BUTTON1_MASK);
	    
	    MAX_X = 323;
		MAX_Y = 613;
	    robot.mouseMove(MAX_X, MAX_Y);
	    robot.mouseRelease(InputEvent.BUTTON1_MASK);
	    
		StringSelection selection3 = new StringSelection(Address_line2);
		Clipboard clipboard3 = Toolkit.getDefaultToolkit().getSystemClipboard();
		clipboard3.setContents(selection3, selection3);
	    
	      robot.keyPress(KeyEvent.VK_CONTROL);
	      robot.keyPress(KeyEvent.VK_V);
	      robot.keyRelease(KeyEvent.VK_V);
	      robot.keyRelease(KeyEvent.VK_CONTROL);
	      
		  //Select Address line State
		    Thread.sleep(2000);
		    MAX_X = 391;
			MAX_Y = 612;
		    robot.mouseMove(MAX_X, MAX_Y);
		    	   
		    robot.mousePress(InputEvent.BUTTON1_MASK);
		    
		    MAX_X = 410;
			MAX_Y = 610;
		    robot.mouseMove(MAX_X, MAX_Y);
		    robot.mouseRelease(InputEvent.BUTTON1_MASK);
		    
			StringSelection selection4 = new StringSelection(Address_line3);
			Clipboard clipboard4 = Toolkit.getDefaultToolkit().getSystemClipboard();
			clipboard1.setContents(selection4, selection4);
		    
		    robot.keyPress(KeyEvent.VK_CONTROL);
		    robot.keyPress(KeyEvent.VK_V);
		    robot.keyRelease(KeyEvent.VK_V);
		    robot.keyRelease(KeyEvent.VK_CONTROL);
		    
		 //Select Address line Zip
		   Thread.sleep(2000);
		   MAX_X = 472;
		   MAX_Y = 611;
		  robot.mouseMove(MAX_X, MAX_Y);
		   	   
		  robot.mousePress(InputEvent.BUTTON1_MASK);
			    
		   MAX_X = 552;
		   MAX_Y = 610;
		  robot.mouseMove(MAX_X, MAX_Y);
		   robot.mouseRelease(InputEvent.BUTTON1_MASK);
		   
		StringSelection selection5 = new StringSelection(Address_line4);
		Clipboard clipboard5 = Toolkit.getDefaultToolkit().getSystemClipboard();
		clipboard5.setContents(selection5, selection5);
		   
		 robot.keyPress(KeyEvent.VK_CONTROL);
		 robot.keyPress(KeyEvent.VK_V);
		 robot.keyRelease(KeyEvent.VK_V);
		 robot.keyRelease(KeyEvent.VK_CONTROL);
		 
		//Click on "statement tab"
		 Thread.sleep(2000);
		 MAX_X = 210;
		 MAX_Y = 484;
		 robot.mouseMove(MAX_X, MAX_Y);
		 	   
		 robot.mousePress(InputEvent.BUTTON1_MASK);
		 robot.mouseRelease(InputEvent.BUTTON1_MASK);
		 
		 Thread.sleep(2000);
		 ImageMatched = false;
		 Robot robot2 = new Robot();
	     String fileName = "C://Driver//HardCopyCheck_Current.jpg";

	     // Define an area of size 500*400 starting at coordinates (10,50)
	     //Rectangle rectArea = new Rectangle(24, 218, 56, 12);
	     Rectangle rectArea = new Rectangle(421, 532, 18, 18);
	     BufferedImage screenFullImage = robot2.createScreenCapture(rectArea);
	     ImageIO.write(screenFullImage, "jpg", new File(fileName));
	     
	   //Compare the image with saved image - 
	     BufferedImage imgA = null; 
	     BufferedImage imgB = null; 
	     
	     try
	     { 
	         File fileA = new File("C://Driver//HardCopyCheck_Current.jpg"); 
	         File fileB = new File("C://Driver//HardCopyCheck.jpg"); 

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
	         if (difference <= 1000){
	      	   System.out.println("Image matched");
	      	   ImageMatched = true;
	         }
	         
	     }
		 //Uncheck Hardcopy
	         if (ImageMatched == true)
	        		 {
						 Thread.sleep(2000);
						 MAX_X = 432;
						 MAX_Y = 540;
						 robot.mouseMove(MAX_X, MAX_Y);
						 	   
						 robot.mousePress(InputEvent.BUTTON1_MASK);
						 robot.mouseRelease(InputEvent.BUTTON1_MASK);
	        		 }
	         
	       //Click on "Bill-To tab"
	    	 Thread.sleep(2000);
	    	 MAX_X = 49;
	    	 MAX_Y = 483;
	    	 robot.mouseMove(MAX_X, MAX_Y);
	    	 	   
	    	 robot.mousePress(InputEvent.BUTTON1_MASK);
	    	 robot.mouseRelease(InputEvent.BUTTON1_MASK);
	    	 
	    	//Press Enter
			 Thread.sleep(1000);
			 robot.keyPress(KeyEvent.VK_ENTER);
			 robot.keyRelease(KeyEvent.VK_ENTER);
			 Thread.sleep(1000);
			 robot.keyPress(KeyEvent.VK_ENTER);
			 robot.keyRelease(KeyEvent.VK_ENTER);     
			 
			 //Enter Memo -  on "Memo Maintenance"
			 Thread.sleep(5000);
			 MAX_X = 190;
			 MAX_Y = 56;
			 robot.mouseMove(MAX_X, MAX_Y);
			 	   
			 robot.mousePress(InputEvent.BUTTON1_MASK);
			 robot.mouseRelease(InputEvent.BUTTON1_MASK);
		
			//Click on permanent memo check box
			 Thread.sleep(2000);
			 MAX_X = 23;
			 MAX_Y = 115;
			 robot.mouseMove(MAX_X, MAX_Y);
			 	   
			 robot.mousePress(InputEvent.BUTTON1_MASK);
			 robot.mouseRelease(InputEvent.BUTTON1_MASK); 
			 
			//Click on  memo  box
			 Thread.sleep(2000);
			 MAX_X = 48;
			 MAX_Y = 150;
			 robot.mouseMove(MAX_X, MAX_Y);
			 	   
			 robot.mousePress(InputEvent.BUTTON1_MASK);
			 robot.mouseRelease(InputEvent.BUTTON1_MASK); 
			 
			 //Paste the Memo in box 
			Thread.sleep(2000); 
		    StringSelection selection6 = new StringSelection(Memo);
			Clipboard clipboard6 = Toolkit.getDefaultToolkit().getSystemClipboard();
			clipboard1.setContents(selection6, selection6);
			robot.keyPress(KeyEvent.VK_CONTROL);
			robot.keyPress(KeyEvent.VK_V);
			robot.keyRelease(KeyEvent.VK_V);
			robot.keyRelease(KeyEvent.VK_CONTROL);
			 

			 //Presee Alt+A to have the Memo added
			 robot.keyPress(KeyEvent.VK_ALT);
			 robot.keyPress(KeyEvent.VK_A);
			 robot.keyRelease(KeyEvent.VK_A);
			 robot.keyRelease(KeyEvent.VK_ALT);
			 
			Thread.sleep(1000);
			 robot.keyPress(KeyEvent.VK_TAB);
			 robot.keyRelease(KeyEvent.VK_TAB); 
			 robot.keyPress(KeyEvent.VK_TAB);
			 robot.keyRelease(KeyEvent.VK_TAB); 		 
			 robot.keyPress(KeyEvent.VK_TAB);
			 robot.keyRelease(KeyEvent.VK_TAB); 
			 robot.keyPress(KeyEvent.VK_TAB);
			 robot.keyRelease(KeyEvent.VK_TAB); 
			 robot.keyPress(KeyEvent.VK_TAB);
			 robot.keyRelease(KeyEvent.VK_TAB); 

			 Thread.sleep(1000);
			 robot.keyPress(KeyEvent.VK_ENTER);
			 robot.keyRelease(KeyEvent.VK_ENTER);
		 
		 
		 
		 //Press Enter
		 Thread.sleep(1000);
		 robot.keyPress(KeyEvent.VK_ENTER);
		 robot.keyRelease(KeyEvent.VK_ENTER);
		 Thread.sleep(1000);
		 robot.keyPress(KeyEvent.VK_ENTER);
		 robot.keyRelease(KeyEvent.VK_ENTER);     
		 
		 
		 //Presee Alt+C to close the window
		 Thread.sleep(2000);
		 robot.keyPress(KeyEvent.VK_ALT);
		 robot.keyPress(KeyEvent.VK_C);
		 robot.keyRelease(KeyEvent.VK_C);
		 robot.keyRelease(KeyEvent.VK_ALT);		 
		 Thread.sleep(1000);
		
	   
		}
		}
	    
	    static void PressPaste() throws Exception{
	    	Robot robot = new Robot();
	    	
	    	robot.mousePress(InputEvent.BUTTON3_MASK);
	    	robot.mouseRelease(InputEvent.BUTTON3_MASK);
	    	Thread.sleep(SleepBtwClicks);	    
		    //System.out.println("MAX_X and MAX_Y"+MAX_X+MAX_Y);
	    	MAX_X = MAX_X + 15;
	    	MAX_Y = MAX_Y + 80;
		   	robot.mouseMove(MAX_X, MAX_Y);
		   	Thread.sleep(SleepBtwClicks);
		   	robot.mousePress(InputEvent.BUTTON1_MASK);
		   	robot.mouseRelease(InputEvent.BUTTON1_MASK);
	    }
	    
	    static void PressPaste2() throws Exception{
	    	Robot robot = new Robot();
	    	
	    	robot.mousePress(InputEvent.BUTTON3_MASK);
	    	robot.mouseRelease(InputEvent.BUTTON3_MASK);
	    	Thread.sleep(SleepBtwClicks);	    
		    //System.out.println("MAX_X and MAX_Y"+MAX_X+MAX_Y);
	    	MAX_X = MAX_X + 10;
	    	MAX_Y = MAX_Y - 190;
		   	robot.mouseMove(MAX_X, MAX_Y);
		   	Thread.sleep(SleepBtwClicks);
		   	robot.mousePress(InputEvent.BUTTON1_MASK);
		   	robot.mouseRelease(InputEvent.BUTTON1_MASK);
	    }
	   
	    static void Copy() throws Exception{
	    	Robot robot = new Robot();
	    	
	    	robot.mousePress(InputEvent.BUTTON3_MASK);
	    	robot.mouseRelease(InputEvent.BUTTON3_MASK);
	    	Thread.sleep(SleepBtwClicks);	    
		    //System.out.println("MAX_X and MAX_Y"+MAX_X+MAX_Y);
	    	MAX_X = MAX_X + 10;
	    	MAX_Y = MAX_Y + 10;
		   	robot.mouseMove(MAX_X, MAX_Y);
		   	Thread.sleep(SleepBtwClicks);
		   	robot.mousePress(InputEvent.BUTTON1_MASK);
		   	robot.mouseRelease(InputEvent.BUTTON1_MASK);
	    }
	    
	    static void Copy2() throws Exception{
	    	Robot robot = new Robot();
	    	
	    	robot.mousePress(InputEvent.BUTTON3_MASK);
	    	robot.mouseRelease(InputEvent.BUTTON3_MASK);
	    	Thread.sleep(SleepBtwClicks);	    
		    //System.out.println("MAX_X and MAX_Y"+MAX_X+MAX_Y);
	    	MAX_X = MAX_X + 40;
	    	MAX_Y = MAX_Y + 70;
	    	
		   	robot.mouseMove(MAX_X, MAX_Y);
		   	Thread.sleep(SleepBtwClicks);
		   	robot.mousePress(InputEvent.BUTTON1_MASK);
		   	robot.mouseRelease(InputEvent.BUTTON1_MASK);
	    }
	    
	}
	
