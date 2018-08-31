package com.deere.Helpers;

import java.util.List;
import java.util.NoSuchElementException;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedCondition;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.FluentWait;
import org.openqa.selenium.support.ui.Wait;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;

import com.google.common.base.Function;


public class WaitFactory extends BaseClass {
	
	
	public static  List<WebElement> WaitForListofElementsToVisible(List<WebElement> elements) throws Exception
	{	List<WebElement>  elem = null;
		    try {
			
				Wait<WebDriver> waitobj = new FluentWait<WebDriver>(BaseClass.wbDriver)
				 
					       .withTimeout(160, TimeUnit.SECONDS)
					 
					       .pollingEvery(2, TimeUnit.SECONDS)
					 
					       .ignoring(NoSuchElementException.class);
				elem =  (List<WebElement>) waitobj.until(ExpectedConditions.visibilityOfAllElements(elements));
			
			} catch (Exception e) {
				// TODO Auto-generated catch block
		//		Assert.fail("Timeout waiting for Element Load Request to complete."); 
			}
		
			return elem;	 
	 }
	
	
public static  WebElement WaitForElementToVisible(WebElement element) throws Exception
{	WebElement elem = null;
	    try {
		
		Wait<WebDriver> waitobj = new FluentWait<WebDriver>(BaseClass.wbDriver)
						.withTimeout(160, TimeUnit.SECONDS)
				       .pollingEvery(1, TimeUnit.SECONDS)
				 
				       .ignoring(NoSuchElementException.class);
			elem = (WebElement) waitobj.until(ExpectedConditions.visibilityOf(element));
		
		} catch (Exception e) {
			// TODO Auto-generated catch block
			Assert.fail("Timeout waiting for Element Load Request to complete."); 
		}
	
		return elem;	 
 }



	
	public static void waitForPageLoaded() { 

		ExpectedCondition<Boolean> expectation = new 
				ExpectedCondition<Boolean>() { 
			public Boolean apply(WebDriver driver) { 
				return ((JavascriptExecutor)wbDriver).executeScript("return document.readyState").equals("complete"); 
			} 
		}; 

		Wait<WebDriver> wait = new WebDriverWait(wbDriver, 30); 
		try { 
			wait.until(expectation); 
		} catch(Throwable error) { 
	//	Assert.fail("Timeout waiting for Page Load Request to complete."); 
		} 
	} 
	
	
	public static WebElement explicitWait(final String strLocator)  
	{  
		Wait<WebDriver> wait = new FluentWait<WebDriver>(wbDriver)  
				.withTimeout(20, TimeUnit.SECONDS)  
				.pollingEvery(5, TimeUnit.SECONDS)  
				.ignoring(NoSuchElementException.class); 

		WebElement element= wait.until(new Function<WebDriver, WebElement>() {  
			public WebElement apply(WebDriver driver) {  
				return wbDriver.findElement(By.id(strLocator));  
			}  
		});  
		return element;  
	}  

	public static WebElement explicitWaitByXpath(final String strLocator)  
	{  
		Wait<WebDriver> wait = new FluentWait<WebDriver>(wbDriver)  
				.withTimeout(20, TimeUnit.SECONDS)  
				.pollingEvery(5, TimeUnit.SECONDS)  
				.ignoring(NoSuchElementException.class); 

		WebElement element= wait.until(new Function<WebDriver, WebElement>() {  
			public WebElement apply(WebDriver driver) {  
				return wbDriver.findElement(By.xpath(strLocator));  
			}  
		});  
		return element;  
	}  
	
	public static void implicitWaitInSeconds(int intWaitTime ){
		wbDriver.manage().timeouts().implicitlyWait(intWaitTime, TimeUnit.SECONDS);
	}
	

	public static WebElement waitForElement(WebElement Wbel) {
		try {
			WebDriverWait wait = new WebDriverWait(wbDriver, 44);
			return wait.until(ExpectedConditions.visibilityOf(Wbel));
		} catch (NoSuchElementException e) {
			throw new NoSuchElementException(e.getMessage());
		}
	}
	
	public static List<WebElement> waitForElements(List<WebElement> elements,
			WebDriver driver) {
		try {
			WebDriverWait wait = new WebDriverWait(wbDriver, 25);
			return wait.until(ExpectedConditions
					.visibilityOfAllElements(elements));
		} catch (NoSuchElementException e) {
			throw new NoSuchElementException(e.getMessage());
		}
	}

	
	public static Alert waitForAlertPresent(WebDriver driver) {
		try {
			WebDriverWait wait = new WebDriverWait(wbDriver, 25);
			return wait.until(ExpectedConditions.alertIsPresent());
		} catch (NoSuchElementException e) {
			throw new NoSuchElementException(e.getMessage());
		}
	}
	
	public static int getRowsFromTable(WebElement Wbel){
		List<WebElement> Rows = Wbel.findElements(By.tagName("tr"));
		return Rows.size();
	}
	
	public static void acceptAlertMessage(WebDriver driver){
		WaitFactory.waitForAlertPresent(wbDriver);
		driver.switchTo().alert().accept();
	}
	
	public static WebElement waitForElementClickable(WebElement Wbel) {
		try {
			WebDriverWait wait = new WebDriverWait(wbDriver, 50);
			return wait.until(ExpectedConditions.elementToBeClickable(Wbel));
		} catch (NoSuchElementException e) {
			throw new NoSuchElementException(e.getMessage());
		}
	}

	
	public static boolean waitForTextIsPresentInElement(WebElement Wbel,
			WebDriver driver, String Strtext) {
		try {
			WebDriverWait wait = new WebDriverWait(wbDriver, 25);
			return wait.until(ExpectedConditions.textToBePresentInElementValue(
					Wbel, Strtext));
		} catch (NoSuchElementException e) {
			throw new NoSuchElementException(e.getMessage());
		}
	}

	
	
	public static WebElement FluentWaitByWebElement(final WebElement strLocator)  
	{  
		Wait<WebDriver> wait = new FluentWait<WebDriver>(BaseClass.wbDriver)  
				.withTimeout(60, TimeUnit.SECONDS)  
				.pollingEvery(2, TimeUnit.SECONDS)  
				.ignoring(NoSuchElementException.class); 

		WebElement element= wait.until(new Function<WebDriver, WebElement>() {  
			public WebElement apply(WebDriver driver) { 
//				System.out.println("Checking for the Element!!");
				if (strLocator != null) {
					System.out.println("Element is found");
				}
				return strLocator;  
			}  
		});  
		return element;  
	}  
	
	public static  boolean WaitForinvisibilityOfElement(WebElement countryApplyFilterButton) throws Exception
	{	boolean elem=false;
		    try {
			
			Wait<WebDriver> waitobj = new FluentWait<WebDriver>(BaseClass.wbDriver)
							.withTimeout(30, TimeUnit.SECONDS)
					       .pollingEvery(3, TimeUnit.SECONDS)
					 
					       .ignoring(NoSuchElementException.class);
				elem = waitobj.until(ExpectedConditions.invisibilityOf((WebElement) countryApplyFilterButton));
			
			} catch (Exception e) {
				// TODO Auto-generated catch block
				Assert.fail("Timeout waiting for Element Load Request to complete."); 
			}
		
			return elem;	 
	 }
	

	}
	
