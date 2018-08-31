
/* 
 * Project    : DealerPath
 * Script     : ValidationFactory
 * Author     : Shrishail Baddi
 * Date       : April.15.2018
 * Last Modified On:
 * Modified By :
 */

package com.deere.Helpers;

import java.io.File;
import java.io.IOException;
import java.security.InvalidParameterException;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.By;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.TimeoutException;
import org.openqa.selenium.WebElement;
import org.testng.Assert;
import org.testng.Reporter;

import com.deere.Helpers.*;


@SuppressWarnings("serial")
class Elementpresence extends Exception {
	Elementpresence(String s) {
		super(s);
	}
}

class Elementenable extends Exception {
	Elementenable(String s) {
		super(s);
	}
}

public class ValidationFactory {
	static WebElement locator = null;

	static List<WebElement> locators = null;
	
	
	public static boolean isElementPresent(By by) {
		try {
			locator = BaseClass.wbDriver.findElement(by);

			if (locator.isDisplayed())
				return true;
			else
				return false;
		} catch (NoSuchElementException e) {
			LogFactory.info("Unable to locate the element- No such Element Found");
			Assert.fail("Unable to locate the element- No such Element Found");
			return false;
		}

		catch (Exception e) {
			e.printStackTrace();
			LogFactory.info("Unable to locate the element- No such Element Found");
			return false;
		}

	}
	
	public static WebElement getElementIfPresent(By by) {
		try {
			locator = BaseClass.wbDriver.findElement(by);

			if (locator.isDisplayed())
				return locator;
			else
				return null;
		} catch (NoSuchElementException e) {
			
			LogFactory.info("Unable to locate the element"+locator+"- No such Element Found");
			return null;
			
		//	Assert.fail("Unable to locate the element- No such Element Found");
	
		}

		catch (Exception e) {
			e.printStackTrace();
			LogFactory.info("Unable to locate the element"+locator+"- No such Element Found");
			return null;
		}

	}
	
	public static boolean isElementPresent(WebElement wbelObj) {
		try {
			
			if (wbelObj.isDisplayed())
				return true;
			else
				return false;
		} catch (NoSuchElementException e) {
			return false;
		}

		catch (Exception e) {
			return false;
		}

	}
	

	public static void elementNotPresent(By by, String strLocator, String strErrorMessage) {
		Boolean present = false;
		try {
			BaseClass.wbDriver.manage().timeouts().implicitlyWait(0, TimeUnit.SECONDS);
			if (!(strLocator).isEmpty())
				throw new noLocatorexceptions(strLocator);
			else {
				locator = BaseClass.wbDriver.findElement(by);
				locator.isDisplayed();
				if (present)
					throw new Elementpresence(strLocator);
			}
		} catch (noLocatorexceptions e) {
			LogFactory.info(strErrorMessage);
			Assert.fail("nmkli0op");
		} catch (Elementpresence e1) {
			screenShot(strLocator);
			Assert.fail(strErrorMessage);
		} catch (NoSuchElementException e) {
			;
		} catch (Exception e) {
			;
		}
	}

	public static void screenShot(String strName) {
		String name1 = System.getProperty("user.dir") + "/test-output/" + strName + ".png";
		try {
			File source = ((TakesScreenshot) BaseClass.wbDriver).getScreenshotAs(OutputType.FILE);
			File destFile = new File(name1);
			FileUtils.copyFile(source, destFile);
			String image = "file:///" + name1;
			Reporter.log("<a href='" + image + "'>ScreenshotLink</a>");
		} catch (IOException e) {
			LogFactory.info("Failed to capture screenshot: " + e.getMessage());
		}

	}

	public static boolean compareTwoStrings(String strActual, String strExpected, String strErrorMessage) {

		boolean result = false;
		try {
			if (strActual.equalsIgnoreCase(strExpected))
				result = true;
			else
				result = false;
			// LogFactory.info(ErrorMessage);
			if (!result)
				throw new Elementpresence("");

		} catch (Elementpresence e) {
			LogFactory.info(strErrorMessage);
			ValidationFactory.screenShot(strExpected);
			Assert.fail(strErrorMessage);
		} catch (Exception e) {
			;
		}
		return result;
	}

	public static boolean compareTwoStringsFalse(String strActual, String strExpected, String strErrorMessage) {

		boolean result = false;
		try {
			if (strActual.equalsIgnoreCase(strExpected))
				result = true;
			else
				result = false;
			if (result)
				throw new Elementpresence("");

		} catch (Elementpresence e) {
			LogFactory.info(strErrorMessage);
			ValidationFactory.screenShot(strExpected);
			Assert.fail(strErrorMessage);
		} catch (Exception e) {
			;
		}
		return result;
	}

	public static boolean compareTwoValues(Integer intActual, Integer intExpected, String strErrorMessage) {
		boolean result = false;
		try {
			if (intActual == intExpected)
				result = true;
			else
				result = false;
			if (!result)
				throw new Elementpresence("");
		} catch (Elementpresence e) {
			LogFactory.info(strErrorMessage);
			Assert.fail(strErrorMessage);
		}

		catch (Exception e) {
			;
		}
		return result;
	}

	public static boolean checkBoolean(Boolean actual, Boolean expected, String strErrorMessage) {
		boolean result = false;
		try {
			if (actual == expected)
				result = true;
			else
				result = false;
			if (!result)
				throw new Elementpresence("");
		} catch (Elementpresence e) {
			LogFactory.info(strErrorMessage);
			Assert.fail(strErrorMessage);
		} catch (Exception e) {
			;
		}

		return result;
	}

	public void checkBooleanValue(int intActual) {
		boolean result = false;
		if (intActual == 0)
			result = true;
		else if (intActual > 0)
			result = false;
	}

	public static void waitForLoadMaskToComplete(String strLocator, int intTimeOut) throws Exception {
		try {
			int i;
			for (i = 0; i <= 20; i++) {
				Thread.sleep(1000);
				WebElement Mask = BaseClass.wbDriver.findElement(By.id(strLocator));
				String status = Mask.getAttribute("class");
				if (status.equalsIgnoreCase("")) {
					Thread.sleep(500);
					break;
				}
			}
		} catch (NoSuchElementException e) {
			LogFactory.info(locator + "");
			screenShot("Screenshotlink");
			Assert.fail(locator + "");
		} catch (Exception e) {
			LogFactory.info(locator + e.getMessage());
			screenShot("Screenshotlink");
			Assert.fail(e.getMessage());
		}
	}

	public static boolean validateButtonDisable(By by) throws Exception {
		boolean value = false;
		try {
			locator = BaseClass.wbDriver.findElement(by);
			value = locator.isEnabled();
			if (value)
				return true;
			throw new Elementenable("");
		} catch (NoSuchElementException e) {
			LogFactory.info("Unable to locate the element- No such Element Found");
			Assert.fail("Unable to locate the element- No such Element Found");
			return false;
		}

		catch (Exception e) {
			e.printStackTrace();
			LogFactory.info("Unable to locate the element- No such Element Found");
			return false;
		}

	}
	


	public static boolean validateButtonEnable(By by) throws Exception {
		boolean value = false;
		try {
			locator = BaseClass.wbDriver.findElement(by);
			value = locator.isEnabled();
			if (value)
				return true;
			throw new Elementenable("");
		} catch (NoSuchElementException e) {
			LogFactory.info("Unable to locate the element- No such Element Found");
		//	Assert.fail("Unable to locate the element- No such Element Found");
			return false;
		}

		catch (Exception e) {
			e.printStackTrace();
			LogFactory.info("Unable to locate the element- No such Element Found");
			return false;
		}

	}
	
/*	public static void validateButtonEnable(By by, String Locator, String errorMessage) throws Exception {
		Boolean value = false;
		try {
			if (!(Locator).isEmpty())
				throw new noLocatorexceptions(Locator);
			else {
				locator = BaseClass.wbDv.findElement(by);
				value = locator.isEnabled();
				if (!value)
					throw new Elementenable("");
			}
		} catch (Elementenable e) {
			LogFactory.info(errorMessage);
			screenShot(errorMessage);
			Assert.fail(errorMessage);
		} catch (noLocatorexceptions e) {
			LogFactory.info(errorMessage + " ID/Xpath not found ");
			screenShot(errorMessage);
			Assert.fail("ID/Xpath not found");
		} catch (NoSuchElementException e) {
			LogFactory.info(errorMessage);
			screenShot(errorMessage);
			Assert.fail(errorMessage);
		} catch (Exception e) {
			e.printStackTrace();
			LogFactory.info(errorMessage);
		}
	}*/

	public boolean checkBooleanNotSame(Boolean actual, Boolean expected, String ErrorMessage) {
		boolean result = false;
		try {
			if (actual != expected)
				result = false;
			else
				result = true;
			if (!result)
				throw new Elementpresence("");
		} catch (Elementpresence e) {
			LogFactory.info(ErrorMessage);
			Assert.fail(ErrorMessage);
		}

		catch (Exception e) {
			;
		}

		return result;
	}
	
	
	
	
	public static List<WebElement> getElementsIfPresent(By by) {
		try {
			locators = BaseClass.wbDriver.findElements(by);

			if (locators.size()>0)
				return locators;
			else
				return null;
		} catch (NoSuchElementException e) {
			
			LogFactory.info("Unable to locate the element list"+locators+"- No such Element Found");
			return null;
			
		//	Assert.fail("Unable to locate the element- No such Element Found");
	
		}

		catch (Exception e) {
			e.printStackTrace();
			LogFactory.info("Unable to locate the element list"+locators+"- No such Element Found");
			return null;
		}

	}
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	

}
