package test;

import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.testng.Assert;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Parameters;
import org.testng.annotations.Test;

import engine.XLUtility;
import io.github.bonigarcia.wdm.WebDriverManager;


public class DataDrivenTest {

	public WebDriver driver;
	public String Path = "src/main/java/data/fake_data2.xlsx";
	public String SheetName = "Set 1";
	
	
	@BeforeClass
	public void setup()
	{
		WebDriverManager.chromedriver().setup();
		ChromeOptions options = new ChromeOptions();
		driver=new ChromeDriver(options);
		
		driver.manage().timeouts().implicitlyWait(5,TimeUnit.SECONDS);
		driver.manage().window().maximize();
	}
	
	
	@Test(dataProvider="LoginData")
	public void loginTest(String username,String password,String expected) throws IOException
	{		
		
		
		
		driver.get("https://admin-demo.nopcommerce.com/login");
		
		WebElement Email = driver.findElement(By.id("Email"));															// By Email
		Email.clear();
		Email.sendKeys(username);
		
		
		WebElement Password = driver.findElement(By.id("Password"));													// By Password
		Password.clear();
		Password.sendKeys(password);
		
		driver.findElement(By.cssSelector("[class='button-1 login-button']")).click(); 									// Login										
		
		String exp_title = "Dashboard / nopCommerce administration";
		String act_title = driver.getTitle();
		
		if(expected.equals("Valid"))																					// If Valid then we would expect a successful test to be matching titles. (Exp vs Actual)
		{	
			if(exp_title.equals(act_title))
			{
				driver.findElement(By.xpath("//*[text()='Logout']")).click();
				Assert.assertTrue(true);
			}
			else
			{
				Assert.assertTrue(false);
			}
		}
		else if(expected.equals("Invalid"))																				// If Invalid then we would expect a successful test to be not matching titles. (Exp vs Actual)
		{
			if(exp_title.equals(act_title))
			{
				driver.findElement(By.xpath("//*[text()='Logout']")).click();
				Assert.assertTrue(false);
			}
			else
			{
				Assert.assertTrue(true);
			}
		}
		
	}
	
	
	@DataProvider(name="LoginData")
	public String [][] getData() throws IOException
	{
		XLUtility xlutil = new XLUtility(Path);
		
		int totalrows=xlutil.getRowCount(SheetName);
		int totalcols=xlutil.getCellCount(SheetName, 1);			
		
		String loginData[][]=new String[totalrows][totalcols];
					
		for(int i=1;i<=totalrows;i++) //1
		{
			for(int j=0;j<totalcols;j++) //0
			{
				loginData[i-1][j]=xlutil.getCellData(SheetName, i, j);
			}			
		}
		
		return loginData;
	}
	
	
	@AfterClass
	void tearDown()
	{
		driver.close();
	}
	
	
	
	
}
