package com.deere.Helpers;

import java.io.IOException;

import org.openqa.selenium.By;
import org.testng.SkipException;

import com.deere.PageFactory.Login_Page_POF;

public class ImpersonateUser extends BaseClass{

	/**
	 * This method is used to impersonate the dealer through his rackfID 
	 * @author shrishail.baddi
	 * @param strRACFID
	 * @createdAt 05-06-2018
	 * @throws IOException
	 * 
	 * @modifyBy shrey.choudhary
	 * @modifyAt 07-06-2018
	 * @return boolean
	 * @throws Throwable 
	 */
/*	public static boolean impersonateUserSuccess(String strRACFID) throws Throwable {
		try {
			BaseClass.errorUserFound = false;
			if (!GenericFactory.isNull(strRACFID)) {
				if (ValidationFactory
						.isElementPresent(WaitFactory.explicitWaitByXpath("//b[text() = 'Admin Links']"))) {
					ValidationFactory.getElementIfPresent(By.xpath("//a[text() = 'Impersonate User']")).click();
					LogFactory.info("Clicked Impersonate User link from left navigation.....");
					WaitFactory.waitForPageLoaded();
					LogFactory.info("Enter User RACF ID.....");
					ValidationFactory.getElementIfPresent(By.xpath("//input [@name = 'inputText_AU']"))
							.sendKeys(strRACFID);

					LogFactory.info("Click on Search Button.....");
					ValidationFactory.getElementIfPresent(By.xpath("//input [ @name = 'Search']")).click();

					// verify if radio button for selecting the user is present or not
					Boolean isPresent = BaseClass.wbDriver
							.findElements(By.xpath("//input[@type='radio' and @name='userDNS']")).size() > 0;
					if (isPresent) {

						LogFactory.info("Select the user.....");
						ValidationFactory.getElementIfPresent(By.xpath("//input[@type='radio' and @name='userDNS']"))
								.click();

						if (ValidationFactory
								.validateButtonEnable(By.xpath("//input[@type='button' and @name='Impersonate']"))) {

							ValidationFactory
									.getElementIfPresent(By.xpath("//input[@type='button' and @name='Impersonate']"))
									.click();
							LogFactory.info("User " + strRACFID + " successfully logged into appllication ");
							
							
						} else {
							LogFactory.error("Unable to locate the button Impersonate/ it is disabled");
							throw new SkipException("Run mode of testcase " + "test" + " is N");
						}
					} else {
						LogFactory.error(strRACFID + " does not exist/No user Found");
						LogFactory.error(BaseClass.wbDriver
								.findElement(By.xpath(".//*[@id='impersonateUserForm']/table/tbody/tr/td[2]/span"))
								.getText());
						BaseClass.errorUserFound = true;
						return false;
					}
				}
			}
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
		return true;
	}*/
	
	
	
	public static boolean impersonateUserSuccess(String strRACFID) throws Throwable {
		try {
			BaseClass.errorUserFound = false;
			if (!GenericFactory.isNull(strRACFID)) {
				if (ValidationFactory
						.isElementPresent(WaitFactory.explicitWaitByXpath("//b[text() = 'Admin Links']"))) {
					
					ValidationFactory.getElementIfPresent(By.xpath(".//*[@id='Analyze_User']")).click();
					LogFactory.info("Clicked Analyze User link from left navigation.....");
					WaitFactory.waitForPageLoaded();
					LogFactory.info("Enter User RACF ID.....");
					ValidationFactory.getElementIfPresent(By.xpath(".//*[@id='analyseUserId']")).sendKeys(strRACFID);
					
					LogFactory.info("Click on Analyze Button.....");
					ValidationFactory.getElementIfPresent(By.xpath(".//*[@id='analyzeUserButton']")).click();
					
						if (ValidationFactory
								.validateButtonEnable(By.xpath(".//*[@id='impersonateUser']"))) {
							

							ValidationFactory
									.getElementIfPresent(By.xpath(".//*[@id='impersonateUser']"))
									.click();
							
							if (ValidationFactory
							.getElementIfPresent(By.xpath("//a[@id='endimpersonatelink']")) != null) {
								
								BaseClass.errorUserFound = false;
								strUserRACFID =strRACFID;
								LogFactory.info("User " + strRACFID + " not able to login due to the error message : DealerPath Application is not available [AccountFlex is down] ");
								return false;
								
							}
							
							LogFactory.info("User " + strRACFID + " successfully logged into appllication ");
							
							
						} else {
							
							BaseClass.errorUserFound = true;
							LogFactory.error("Unable to locate the button Impersonate/ it is disabled");
							throw new SkipException("Run mode of testcase " + "test" + " is N");
							
							
						}
				}
			}
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
		return true;
	}
	
}