package com.deere.Helpers;

import java.io.IOException;
import java.util.Iterator;
import java.util.List;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import com.deere.PageFactory.Login_Page_POF;
import com.deers.alerts_WCM.Alert_WCM_POF;
import com.steadystate.css.util.ThrowCssExceptionErrorHandler;

public class Invoke_WCM extends BaseClass {
	
	
	WebDriver driver;
	/**
	 * This method is the first step of DealerPath suite which sets user
	 * credentials, initiate drivers and page elements
	 * 
	 * @author shrishail.baddi
	 * @createdAt 07-06-2018
	 * @throws IOException
	 * @throws Exception
	 * @modifyBy shrey.choudhary
	 * @modifyAt
	 */
	@BeforeClass
	public void systemConfigSetup() throws IOException, Exception {
		try {

				BrowserFactory.initiateDriver();
				initPageElements();
				
				

				
			} catch (Exception e) {
				LogFactory.info(e.getMessage());
			} catch (Throwable e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	/**
	 * This method is use to invoke admin's login credentials then go to impersonate
	 * the dealer
	 * 
	 * @author shrey.choudhary
	 * @createdAt 07-06-2018
	 * @throws IOException
	 * @throws Exceptionss88593
	 * @modifyBy
	 * @modifyAt
	 * @throws Throwable
	 */
	@Test(priority=0)
	public static void invokeUserCredentials() throws Throwable {
		
		System.out.println("Verify Valid Login");
		loginPageFactory.setCredentials(strUserName, strPassword);

		if (loginPageFactory.verifyUserLogin()) {
			
			System.out.println("Log in sucessfully");
			
			System.out.println("Navigating to WCM page");
			loginPageFactory.navigateToWCM();
			
			
		
			
		}
		else {
			
			System.out.println("Login for"+BaseClass.strUserName+"Failed");
		}
	}
	
	
	//read all alerts
	
	@Test(priority=1)
	public static void moveToAlert() throws Throwable{
		try {
		System.out.println("***Test case for Alert WCM content verification***");
		
				
		List<WebElement> allLanguages=Alert_WCM_POF.identifyAlllanguages(alertRegion);
		
		Iterator<WebElement> iter = allLanguages.iterator();
		Alert_WCM_POF.createWCMExcel();
		
		
		
		while(iter.hasNext()) {
		    
			WebElement we = iter.next();
			
			System.out.println("Fetching WCM content of alerts for Region:"+we.getText());
			
			Alert_WCM_POF.moveToAlertSection(we);
			
			Alert_WCM_POF.readWCMAlertContent();
	         
				}

		
		
		}
		catch(Exception e)
		{
			System.out.println("Error while navigating to alert section::"+ e.getMessage().toString());
			
		}
	} 
	
	
	
	//read all announcments
	
	/*@Test(priority=2)
	public static void readWCMAnnouncementsContent() throws Throwable
	{
		
		
		try {
			
			System.out.println("navigating inside announcements");
			
			Alert_WCM_POF.moveToAnnouncementSection(alertRegion);
			
			Alert_WCM_POF.readWCMAlertContent();
			
			}
		
		
		catch(Exception e)
		{
			
			System.out.println("Error while navigating to announcement section "+e.getMessage().toString());
			
		}
	}*/
	
	
	/*@AfterClass
	public void closeDriver() {
		
		
		BaseClass.wbDriver.quit();
		System.out.println("all windows closed sucessfully");
	}*/
}