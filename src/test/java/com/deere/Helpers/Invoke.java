package com.deere.Helpers;

import java.io.IOException;

import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import com.deere.PageFactory.Login_Page_POF;

public class Invoke extends BaseClass {
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
				//ExcelFactory.readWCMContentData();
				//ExcelFactory.setCredentials();
				BrowserFactory.initiateDriver();
				initPageElements();
				//GenericFactory.createHeaderSection("Login Page");
				
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
	@Test
	public static void invokeUserCredentials() throws Throwable {
		LogFactory.beginTestCase("Verify Valid Login");
		System.out.println("Verify Valid Login");
		loginPageFactory.setCredentials(strUserName, strPassword);

		if (loginPageFactory.verifyUserLogin()) {
			
			System.out.println("Log in sucessfully");
			//GenericFactory.utilityMenuAdminClick();
			//IterationTest.testNgSuite();
		}
		else {
			/*ReportFactory.reporterOutput("Login & Homepage", "Verify user login",
					BaseClass.strUserName + "\n" + "********", "User should redirect to homepage", "User unable to login due to unknown reason", "FAIL");*/

			System.out.println("Login for"+BaseClass.strUserName+"Failed");
		}
	}
	@AfterClass
	public void closeDriver() {
		BaseClass.wbDriver.close();
	}
}