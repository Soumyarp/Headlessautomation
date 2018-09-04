package Utility;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.htmlunit.HtmlUnitDriver;

public class Headless {

	public static void main(String[] args) {
	//	HtmlUnitDriver driver = new HtmlUnitDriver();
		ChromeOptions option = new ChromeOptions();
	//	option.addArguments("window-size=1400,800");
		option.addArguments("headless");
	System.setProperty("webdriver.chrome.driver", System.getProperty("user.dir")+"/ChromeDriver/chromedriver.exe");
		WebDriver driver = new ChromeDriver(option);
			
	//driver.get("https://qa-org.securtime.in");
		driver.get("https://securax.securtime.in");
	System.out.println(driver.getTitle());
	driver.findElement(By.xpath("//input[@type='email']")).clear();
	driver.findElement(By.xpath("//input[@type='email']")).sendKeys("ranjan@securax.in");
	
	driver.findElement(By.xpath("//input[@type='password']")).clear();
	driver.findElement(By.xpath("//input[@type='password']")).sendKeys("Securax@123");
	
	driver.findElement(By.xpath("//button[@type='submit']")).click();
	
	System.out.println(driver.getTitle()+"-------"+driver.getCurrentUrl());
//	System.out.println(driver.getCurrentUrl());
	
	
	
	

	}
	
		
	}


