package WeChart;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;

public class Login {
	
	static String[] URL= new String[10];
	static String[] username =new String[10];
	static String[] password=new String[10];
	static String[] result= new String[10];
	static String[] comment= new String[10];
	
	
	static File file = new File("C:\\Users\\Aditya\\Documents\\Selenium\\TestDataInput.xlsx");
	public static void main(String[] args) throws IOException {
		
	
		ReadExcelFile();
		
		int count=getInputCount();
		for(int j=0;j<count;j++)
		{
			System.setProperty("webdriver.gecko.driver","C:\\geckodriver-v0.19.0-win64\\geckodriver.exe");
			WebDriver driver = new FirefoxDriver();
			//wait for page to load
			driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
			
			//go to url
			driver.get(URL[j]);
	
			Boolean isURLOpened=false;
			isURLOpened = driver.findElements(By.id("email")).size()>0;
			if(isURLOpened)
			{
				//Maximize the window
				driver.manage().window().maximize();
				
				//fetch username and password
			
				driver.findElement(By.id("email")).sendKeys(username[j]);
				driver.findElement(By.id("password")).sendKeys(password[j]);
				driver.findElement(By.cssSelector("button[type='submit']")).click(); 
				Boolean ispresent=null;
				ispresent = driver.findElements(By.id("role")).size()>0;
				if(ispresent)
				{
					comment[j] = "Success";
				WriteExcelFile("Test case passed",comment[j],j+1);
				}
				else
				{
					comment[j] = "Username or password is invalid.";
					WriteExcelFile("Test case failed",comment[j], j+1);	
				}
				}
			else
			{
				comment[j] = "URL is invalid";
				WriteExcelFile("Test case failed",comment[j], j+1);
			}
			
			driver.close();
					
		}
			
		}
	public static int getInputCount() throws IOException
	{
        FileInputStream iFile = new FileInputStream(file);
        XSSFWorkbook wb = new XSSFWorkbook(iFile);  
        XSSFSheet sheet  = wb.getSheet("Logindata");
        int rowCount = sheet.getLastRowNum();
        return rowCount;
    }
	

	public static void ReadExcelFile()
	{
	try
    {

        FileInputStream iFile = new FileInputStream(file);
        XSSFWorkbook wb = new XSSFWorkbook(iFile);  
        XSSFSheet sheet  = wb.getSheet("Logindata");
        int rowCount = sheet.getLastRowNum();
        System.out.println("the no of rows are : " + rowCount);
        
        for (int row=1; row<=rowCount; row++)
        {
        	
            username[row-1] = sheet.getRow(row).getCell(1).getStringCellValue();
            password[row-1] = sheet.getRow(row).getCell(2).getStringCellValue();
            URL[row-1]=sheet.getRow(row).getCell(0).getStringCellValue();
        }

        iFile.close();
    }    
    catch (IOException e)
    {
        e.printStackTrace();
    }
	}
	
	public static void WriteExcelFile(String Result,String comment, int k) throws IOException
	 
	{
		 try {
			 
			FileInputStream iFile = new FileInputStream(file);
			XSSFWorkbook wb = new XSSFWorkbook(iFile);
			XSSFSheet sheet  = wb.getSheet("Logindata");
			System.out.println("writing class");
			
		    XSSFRow row= sheet.getRow(k);
			XSSFCell resultcell =row.createCell(3);
			XSSFCell commentcell =row.createCell(4);
			resultcell.setCellValue(Result);
			commentcell.setCellValue(comment);
			
	        FileOutputStream oFile =new FileOutputStream(file);
			wb.write(oFile);
			oFile.close();
			
		} catch (FileNotFoundException e) {
			
			e.printStackTrace();
		}
		
		 }
		 
	
	
}
	
	
	
	




