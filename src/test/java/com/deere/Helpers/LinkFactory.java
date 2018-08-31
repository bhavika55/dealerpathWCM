
/* 
 * Project    : DealerPath
 * Script     : LinkFactory
 * Author     : Shrishail Baddi
 * Date       : April.11.2018
 * Last Modified On:
 * Modified By :
 */





package com.deere.Helpers;

/* 
 * Project    : Dealer Path
 * Script     : LinkHelper
 * Author     : 
 * Date       : 
 */

import java.io.File;
import java.util.NoSuchElementException;
import java.util.concurrent.TimeUnit;
import org.apache.commons.io.FileUtils;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.By;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedCondition;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.FluentWait;
import org.openqa.selenium.support.ui.Wait;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;

import com.google.common.base.Function;


class noLocatorexceptions  extends Exception
{
	noLocatorexceptions(String s)
	{super(s);}
}

public class LinkFactory {
	
	public static WebDriver web;	
	/**
	 * This Method is for clicking a link By Xpath.
	 * @param : locator
	 * @exception Exception: If any exception is found
	 * @exception NoSuchElementException : If element is not found
	 * 
	 * */
	public static void clickLinkByXpath(String strlocator, String strObjLocator) throws Exception {
		LogFactory.info("Executing click link event");
		try{   
			if(!(strlocator).isEmpty())   
			{throw  new noLocatorexceptions(strlocator);}
			else{
				visibilityOfElementLocated(strlocator);
				WebElement elems=BaseClass.wbDriver.findElement(By.xpath(strlocator));//Menu Item
				Actions builder = new Actions(BaseClass.wbDriver);
				builder.moveToElement(elems).build().perform();
				Thread.sleep(1000);
				elems.click();
			} 
		}
		catch(NoSuchElementException e) {
			LogFactory.info(strObjLocator  +  " link not found");
			e.printStackTrace();
			ValidationFactory.screenShot(strObjLocator);
			Assert.fail(strObjLocator+" link is not found ");
		} 
		catch(noLocatorexceptions e) {
			LogFactory.info(strObjLocator  +  " Xpath locator not found");
			e.printStackTrace();
			Assert.fail(strObjLocator+" Xpath Locator  not found ");
		} 
		catch(Exception e) {
			LogFactory.info(strObjLocator  +  " link not found");
			e.printStackTrace();
			Assert.fail(strObjLocator+" link is not found ");
		}
	}
	/**
	 * This Method is for clicking a link By LinkText.
	 * @author  
	 * 
	 * */
	public static boolean clickByLinkText(String strLinkText) {

		try{
			BaseClass.wbDriver.findElement(By.linkText(strLinkText)).click();
			LogFactory.info("Clicked on " + strLinkText + " element");
			System.out.println("Clicked on " + strLinkText + " element");
			return true;
		} 
		catch(NoSuchElementException e) {
			LogFactory.info(strLinkText  +  " link not found");
			System.out.println(strLinkText  +  " link not found");
			return false;
		} 
		catch (Exception e) {
			LogFactory.info("clickByLinkText: "+e.getMessage());
			System.out.println("clickByLinkText: "+e.getMessage());
			return false;
		}
	}
	/**
	 * This Method is for clicking a link By id.
	 * @param : locator
	 * @exception Exception: If any exception is found
	 * @exception NoSuchElementException : If element is not found
	 * 
	 * */
	public static void clickLinkById(String strLocator, String strObjLocator) {
		try
		{ if(!(strLocator).isEmpty())   
		{throw  new noLocatorexceptions(strLocator);}
		else{

			visibilityOfElementLocated(strLocator);
			WebElement elems=BaseClass.wbDriver.findElement(By.id(strLocator));//Menu Item
			if(elems.isEnabled()){
				Actions builder = new Actions(BaseClass.wbDriver);
				builder.moveToElement(elems).build().perform();
				Thread.sleep(1000);
				elems.click();
				
			}
		}
		}
		catch(NoSuchElementException e) {
			LogFactory.info(strObjLocator  +  " link not found");
			e.printStackTrace();
			ValidationFactory.screenShot(strObjLocator);
			Assert.fail(strObjLocator+" link is not found ");
		} 
		catch(noLocatorexceptions e) {
			LogFactory.info(strObjLocator  +  " Xpath locator not found");
			e.printStackTrace();
			ValidationFactory.screenShot(strObjLocator);
			Assert.fail(strObjLocator+" Xpath Locator  not found ");
		} 
		catch(Exception e) {
			LogFactory.info(strObjLocator  +  " link not found");
			e.printStackTrace();
		}
	}
	/**
	 * This Method is for checking if link is enabled
	 * @param locator
	 * @return
	 */
	public static String assertEnablelinkbyid(String strLocator){
		LogFactory.info("Executing assert enable link event ");
		try {

			WebElement link = BaseClass.wbDriver.findElement(By.id(strLocator)); 
			if(link.isEnabled()){
				BaseClass.wbDriver.findElement(By.id(strLocator));
			}
		} 
		catch (NoSuchElementException e){
			LogFactory.info("Link is not enabled");
			e.printStackTrace();
		} 
		catch (Exception e) {
			return "Fail:Link is not enabled";
		}
		return "Pass";}

	/**
	 * This Method is for checking if link is enabled
	 * @param locator
	 * @return
	 */
	public static String assertEnableLinkByXpath(String strLocator){
		LogFactory.info("Executing assert enable link event ");
		try {

			WebElement link = BaseClass.wbDriver.findElement(By.xpath(strLocator)); 
			if(link.isEnabled()){
				BaseClass.wbDriver.findElement(By.xpath(strLocator));
			}
		} 
		catch (NoSuchElementException e){
			LogFactory.info("Link is not enabled");
			e.printStackTrace();
		} 
		catch (Exception e) {
			return "Fail:Link is not enabled";
		}
		return "Pass";}

	/**
	 * This Method is for checking if link is disabled
	 * @param : locator
	 */
	public static String assertDisableLink(String strLocator){
		LogFactory.info("Executing assert disable link event ");
		try {
			WebElement link = BaseClass.wbDriver.findElement(By.id(strLocator));
			if(link.isDisplayed()){
				BaseClass.wbDriver.findElement(By.id(strLocator)).click();
			}
		} 
		catch (NoSuchElementException e){
			LogFactory.info("Link is not disabled");
			e.printStackTrace();
		} 
		catch (Exception e) {  
			return "Fail:Link is Enabled";
		}
		return "Pass";
	}

	/**
	 *This Method for finding whether link is present
	 * @param locator
	 * @return
	 */

	public static ExpectedCondition<WebElement> visibilityOfElementLocated(final String strLocator) {
		return new ExpectedCondition<WebElement>() {
			public WebElement apply(WebDriver driver) {
				WebElement toReturn = null;
				try {
					toReturn = driver.findElement(By.id(strLocator));
					if(toReturn.isDisplayed()) {
						return toReturn;
					}
				} 
				catch (Exception e) {
					e.printStackTrace();
					Assert.fail("Element is not found on the web page");
				}
				return toReturn;
			}
		};
	}

	/**
	 * method for link present
	 * @param driver
	 * @return
	 */
	public static String elementPresent(String strLocator,String strObjLocator,String strErrorMsg) {
		LogFactory.info("Checking for an element present");
		try { 
			WebElement abc=BaseClass.wbDriver.findElement(By.id(strLocator));
			if(abc.isDisplayed()){
				LogFactory.info(strObjLocator+"Element is present ");

			}
		}
		catch (NoSuchElementException e) {
			LogFactory.info(strObjLocator+"element not present ");
			//LogFactory.info(ErrorMsg);
			e.printStackTrace();
		}
		catch (Exception e) {
			LogFactory.info(strErrorMsg);
			LogFactory.info(strObjLocator+"element not present ");
			Assert.fail(strObjLocator+"element not present");
		

		}


		return strLocator;
	}

	/**
	 * This Method is for clicking a link By Xpath.
	 * @param : locator
	 * @param : objLocator
	 * @exception Exception: If any exception is found
	 * @exception NoSuchElementException : If element is not found
	 * 
	 * */
	public static boolean click(String strLinkLocator, String strObjLocator) {

		String path_locator= strLinkLocator;
		String locator= path_locator.split(",")[1];
		String identifier=path_locator.split(",")[0];

		LogFactory.info( " Clicking on " + strObjLocator + " link");

		try{
			if(identifier.equalsIgnoreCase("xpath")){ 

				BaseClass.wbDriver.findElement(By.xpath(locator)).click();
				LogFactory.info("Clicked on " + strObjLocator + " button");
				return true;

			} else if (identifier.equalsIgnoreCase("id")) {
				BaseClass.wbDriver.findElement(By.id(locator)).click();
				LogFactory.info("Clicked on " + strObjLocator + " button");
				return true; 
			} else { 
				LogFactory.info("OR is returning null");
				return true; 
			}
		} 
		catch(NoSuchElementException e) {
			LogFactory.info(strObjLocator  +  " link not found");
			e.printStackTrace();
			return false;
		} 
		catch (Exception e) {
			e.printStackTrace();
			return false;
		}
	}

	public static String linkverifyTextNotPresentbyid(String strActual, String strExpected) {
		LogFactory.info("Executing verify text");

		try {
			String Actual = BaseClass.wbDriver.findElement(By.id(strActual)).getText();
			String Expected = strExpected;
			if(Actual!=Expected) {
				LogFactory.info("Actual text is not same as Expected text");
				LogFactory.info("Actual text is "+ Actual);
				LogFactory.info("Expected text is "+ Expected);
				return "Pass";
			} else {
				LogFactory.info("Actual text is same as Expected text");
				LogFactory.info("Actual text is "+ Actual);
				LogFactory.info("Expected text is "+ Expected);
				return "Fail - Actual text is same as Expected text";
			}
		}
		catch(Exception e) {
			e.printStackTrace();    
			LogFactory.info("Actual text is not same as Expected text");
			return "Fail - text not matched ";
		}  
	}
	public static String getTextForLinkbyID(String strLocator){
		String returnValue= null;
		try{
			returnValue = BaseClass.wbDriver.findElement(By.id(strLocator)).getText();
		}catch (NoSuchElementException e){
			e.printStackTrace();
			Assert.fail("Link Text not found");
		} catch (Exception e) {
		}
		return  returnValue ;
	}

	






	public static String assertEnablelinkbyid1(String strLocator){
		LogFactory.info("Executing assert enable link event ");
		try {

			WebElement link = BaseClass.wbDriver.findElement(By.id(strLocator)); 
			if(link.isEnabled()){
				BaseClass.wbDriver.findElement(By.id(strLocator));
			}
		} 
		catch (NoSuchElementException e){
			LogFactory.info("Link is not enabled");
			e.printStackTrace();
		} 
		catch (Exception e) {
			return "Fail:Link is not enabled";
		}
		return "Pass";}  



	public static WebElement fluentWait(final By byLocator){
		Wait<WebDriver> wait = new FluentWait<WebDriver>(BaseClass.wbDriver)
				.withTimeout(30, TimeUnit.SECONDS)
				.pollingEvery(5, TimeUnit.SECONDS)
				.ignoring(NoSuchElementException.class);

		WebElement foo = wait.until(
				new Function<WebDriver, WebElement>() {
					public WebElement apply(WebDriver driver) {
						return driver.findElement(byLocator);
					}
				}
				);
		return  foo;              }     ;


		public static boolean waitForPageFullyLoaded(String strLocator, int intTimeoutMs) throws InterruptedException {
			int previous;
			int current = 0;
			int timeSliceMs = 1000;
			do {
				previous = current;
				Thread.sleep(timeSliceMs);
				intTimeoutMs -= timeSliceMs;
				current = BaseClass.wbDriver.findElements(By.xpath(strLocator)).size();
			} while(current > previous && intTimeoutMs > 0);
			if(intTimeoutMs > 0) {
				return true;
			}
			return false;
		} 
		/**
		 * This Method is for checking if link is enabled
		 * @param locator
		 * @return
		 */
		 public static String assertEnablelink(String strLocator){
			 LogFactory.info("Executing assert enable link event ");
			 try {
				 WebElement link = BaseClass.wbDriver.findElement(By.id(strLocator)); 
				 if(link.isEnabled()){
					 BaseClass.wbDriver.findElement(By.id(strLocator));
				 }
			 } 
			 catch (NoSuchElementException e){
				 LogFactory.info("Link is not enabled");
				 e.printStackTrace();
			 } 
			 catch (Exception e) {
				 return "Fail:Link is not enabled";
			 }
			 return "Pass";}
		 
		 public static String WAITforlink(String strLocator){
			 //	                 		   LogFactory.info("Checking for an element present");
			 WebDriverWait wait = new WebDriverWait(BaseClass.wbDriver, 30);
			 try { 
				 WebElement elem=wait.until(ExpectedConditions.visibilityOfElementLocated(By.id(strLocator)));
				 if(elem.isEnabled()){
					 LogFactory.info("link is present "); 
				 }
			 }
			 catch (NoSuchElementException e) {
				 LogFactory.info("link not present ");
				 e.printStackTrace();
			 }
			 catch (Exception e) {
				 Assert.fail("link not present");

			 }

			 return strLocator; 

		 }  

		 public static boolean waitForPageFullyLoadedID(String strLocator, int intTimeoutMs) throws InterruptedException {
			 int previous;
			 int current = 0;
			 int timeSliceMs = 1000;
			 try{ do {
				 previous = current;
				 Thread.sleep(timeSliceMs);
				 intTimeoutMs -= timeSliceMs;
				 current = BaseClass.wbDriver.findElements(By.id(strLocator)).size();
			 } while(current > previous && intTimeoutMs > 0);
			 }catch(NoSuchElementException e){ 
				 LogFactory.info(strLocator+"element not present ");
				 ValidationFactory.screenShot(strLocator);
				 e.printStackTrace();
			 }
			 if(intTimeoutMs > 0) {
				 return true;
			 }
			 return false;
		 }
		 public static WebElement findElementWithTimeoutWait(By by, long longTimeOut) throws InterruptedException { 
			 WebElement e = null; 
			 long elapsedTime = 0; 
			 while(elapsedTime < longTimeOut) { 
				 try { 
					// ButtonHelper.LoadingMaskAndServerMessageValidationFactory();
					 elapsedTime++; 
					 Thread.sleep(1000); 
					 e = BaseClass.wbDriver.findElement(by); 

					 break; 
				 } catch (NoSuchElementException nse) { 
					 Assert.fail("Waited for"+longTimeOut);
				 } 
			 } 
			 return e; 
		 }  	


		 public static boolean isLinkPresentById(String strLocator, String strObjLocator) {
			 try {
				 boolean flag= BaseClass.wbDriver.findElement(By.id(strLocator)).isDisplayed();
				 return flag;
			 }
			 catch (NoSuchElementException e)
			 {
				 LogFactory.info("element not present "+strObjLocator);
				 return false;

			 }
			 catch (Exception e)
			 {
				 LogFactory.info("element not present "+strObjLocator);
				 return false;
			 }
		 }
		 public static boolean isLinkntPresentById(String strLocator, String strObjLocator) {
			 try {
				 BaseClass.wbDriver.manage().timeouts().implicitlyWait(0, TimeUnit.SECONDS);

				 boolean flag= BaseClass.wbDriver.findElement(By.id(strLocator)).isDisplayed();
				 return flag;
			 }
			 catch (NoSuchElementException e)
			 {
				 LogFactory.info("element not present "+strObjLocator);
				 return false;

			 }
			 catch (Exception e)
			 {
				 LogFactory.info("element not present "+strObjLocator);
				 return false;
			 }
		 }
		 public static boolean isLinkNotPresentById(String strLocator, String strObjLocator) {
			 try {
				 BaseClass.wbDriver.manage().timeouts().implicitlyWait(0, TimeUnit.SECONDS);
				 

				 boolean flag= BaseClass.wbDriver.findElement(By.id(strLocator)).isDisplayed();
				 
				 return flag;
				 
				 
			 }
			 catch (NoSuchElementException e)
			 {
				 BaseClass.wbDriver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
				 LogFactory.info("element not present "+strObjLocator);
				 return false;

			 }
			 catch (Exception e)
			 {
				 BaseClass.wbDriver.manage().timeouts().implicitlyWait(0, TimeUnit.SECONDS);
				 LogFactory.info("element not present "+strObjLocator);
				 return false;
			 }
			 
		 }
		 
		 public static boolean isLinkNotPresentByXpath(String strLocator, String strObjLocator) {
			 try {
				 BaseClass.wbDriver.manage().timeouts().implicitlyWait(0, TimeUnit.SECONDS);
				 

				 boolean flag= BaseClass.wbDriver.findElement(By.xpath(strLocator)).isDisplayed();
				 
				 return flag;
				 
				 
			 }
			 catch (NoSuchElementException e)
			 {
				 BaseClass.wbDriver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
				 LogFactory.info("element not present "+strObjLocator);
				 return false;

			 }
			 catch (Exception e)
			 {
				 BaseClass.wbDriver.manage().timeouts().implicitlyWait(0, TimeUnit.SECONDS);
				 LogFactory.info("element not present "+strObjLocator);
				 return false;
			 }
			 
		 }
		 public static  void WaitForLoadMaskToComplete(String strLocator, int intTimeOut) throws Exception {
			 int i;
			 try{
				 for(i=0;i<=20;i++)
				 {
					 Thread.sleep(1000);
					 WebElement Mask =BaseClass.wbDriver.findElement(By.id(strLocator));
					 String status=Mask.getAttribute("class");
					 if(status.equalsIgnoreCase(""))
					 {
						 Thread.sleep(500);	
						 break;
					 }
				 } 
			 }catch(NoSuchElementException e) {
				 LogFactory.info(strLocator+"");
				 ValidationFactory.screenShot("Screenshotlink");
			 } 
			 catch(Exception e) {
				 LogFactory.info(strLocator+e.getMessage());
				 ValidationFactory.screenShot("Screenshotlink");


			 }}}