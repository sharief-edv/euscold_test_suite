package com.euscold.base;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.sql.Statement;
import java.util.Properties;

import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.BeforeSuite;

import com.ibm.as400.access.AS400JDBCDataSource;

public class DBConnection extends htmlLayout{
	protected WebDriver driver;
	protected String Data = null;
	protected String No_data =null;
    protected Logger log = Logger.getLogger("eUSColdlogger");
	protected Statement stmt;
	protected AS400JDBCDataSource datasource = new AS400JDBCDataSource("DEVDB2.USCOLD.COM");
	protected FileOutputStream fileout;
	protected HSSFWorkbook WB = new HSSFWorkbook();
	protected HSSFSheet sheet = null;
	protected HSSFRow rowhead=null;
	HSSFRow row=null;
	protected int T_qty;
	protected int P_qty;
	protected int H_qty;
	protected int R_qty;
	protected int O_qty;
    		
	static String Logpath = null;
	protected static String Cust_Number = null;    
    protected static String UI_from_date = null;
    protected static String UI_to_date= null;
    protected static String Path = null;
    protected static String test_results = null;
    static String driver_path = null;
    static String sit_url = null;
    static String username = null;
    static String password = null;
    
    
    
    @BeforeSuite
    public void beforesuite() {
    	try {
        	PropertyConfigurator.configure(System.getProperty("user.dir")+"\\src\\test\\resources\\properties\\log4j.properties");
        	
	    	File file = new File(System.getProperty("user.dir")+"\\src\\test\\resources\\properties\\Config.properties");
			FileInputStream fileInput = new FileInputStream(file);
			Properties config = new Properties();
			config.load(fileInput);			
			fileInput.close();
			Logpath = config.getProperty("log.path");
			Cust_Number = config.getProperty("Cust_No");
		    UI_from_date = config.getProperty("Sit_from_date");
		    UI_to_date = config.getProperty("Sit_to_date");
		    Path = config.getProperty("excel_path");
		    test_results = config.getProperty("testresults_path");
		    driver_path = config.getProperty("driverpath");
		    sit_url = config.getProperty("url");
		    username = config.getProperty("username");
		    password = config.getProperty("password");
		   // System.setProperty("logFilename", test_results+"Data_Comparision_Report");
        }catch (Exception e) {
			e.printStackTrace();
		}
    }
    
    void user_details() throws Exception {
    	driver.findElement(By.id("userId")).sendKeys(username);
		driver.findElement(By.id("password")).sendKeys(password);
		driver.findElement(By.xpath(".//*[@id='loginBoxNew']/input[5]")).click();
		Thread.sleep(5000);
    }
    
    public void login() throws Exception {
    	System.setProperty("webdriver.chrome.driver", driver_path);
		driver = new ChromeDriver();
		driver.get(sit_url);
		driver.manage().window().maximize(); 
		Thread.sleep(3000);
		user_details();
		String customer = driver.findElement(By.xpath("//span[1][@class='bld']")).getText();
		if(customer.contains("PERDUE FOODS, INC.")) {
			System.out.println("Customer is already selected");
		}else {
			driver.findElement(By.xpath("//a[1][text()='My Profile']")).click();
			Thread.sleep(2000);
			new Select(driver.findElement(By.xpath("//*[@id='userProfile_customerCompanyId']"))).selectByVisibleText("PERDUE FOODS, INC.");
			Thread.sleep(5000);
			driver.findElement(By.xpath("//a[@class='btnGreen']/span")).click();
			Thread.sleep(5000);
			driver.findElement(By.linkText("Logout")).click();
			Thread.sleep(4000);
			user_details();
		}
    }
}
