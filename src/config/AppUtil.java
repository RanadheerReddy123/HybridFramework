package config;

import java.io.FileInputStream;
import java.util.Properties;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.Reporter;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.BeforeSuite;

public class AppUtil {
	public static WebDriver driver;
	public static Properties config;
	@BeforeSuite
	public static void setUp()throws Throwable
	{
		config= new Properties();
		config.load(new FileInputStream("C:/RR/eclipse-workspace/Hybrid_FrameWork/PropertyFiles/Environment.properties"));
		if(config.getProperty("Browser").equalsIgnoreCase("chrome"))
		{
			System.out.println("Executing on Chrome browser");
			driver = new ChromeDriver();
			driver.manage().window().maximize();
			driver.get(config.getProperty("Url"));
			driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		}
		else if(config.getProperty("Browser").equalsIgnoreCase("firefox"))
		{
			System.out.println("Executing on Firefox browser");
			driver = new FirefoxDriver();
			driver.manage().window().maximize();
			driver.get(config.getProperty("Url"));
			driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		}
		else
		{
			Reporter.log("Browser value is Not Matching",true);
		}

	}
	@AfterSuite
	public static void tearDown()
	{
		driver.quit();
	}
}
