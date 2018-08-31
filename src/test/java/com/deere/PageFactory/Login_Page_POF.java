
/* 
 * Project    : DealerPath
 * Script     : Login_Page_POF
 * Author     : Shrishail Baddi
 * Date       : April.14.2018
 * Last Modified On:
 * Modified By :
 */

package com.deere.PageFactory;

import java.util.Iterator;
import java.util.Set;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.How;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import com.deere.Helpers.BaseClass;
import com.deere.Helpers.GenericFactory;
import com.deere.Helpers.LogFactory;
import com.deere.Helpers.ReportFactory;
import com.deere.Helpers.ValidationFactory;

public class Login_Page_POF extends BaseClass {

	static WebDriver locDriver;

	public Login_Page_POF(WebDriver driver) {

		this.locDriver = driver;

	}

	@FindBy(how = How.ID_OR_NAME, using = "username")
	static WebElement wbelTxtUserName;
	@FindBy(how = How.NAME, using = "PASSWORD")
	static WebElement wbelTxtPasswordName;
	@FindBy(how = How.NAME, using = "login")
	static WebElement wbelBtnSignin;
	
	@FindBy(how = How.XPATH, using = "//a[contains (.,'WCM')]")
	static WebElement wcmLink;
	
	

	public static boolean setCredentials(String mStrUserName, String mStrPassword) throws Throwable {

		try {
			
			String strFlagUserName = "Fail";
			wbelTxtUserName.sendKeys(mStrUserName);

			if ((wbelTxtUserName.getText()) != "") {
				strFlagUserName = "Pass";
				LogFactory.info(" Username... " + wbelTxtUserName.getText());
			}
			String strFlagPassword = "Fail";
			wbelTxtPasswordName.sendKeys(mStrPassword);

			if ((wbelTxtPasswordName.getText() != "")) {
				strFlagPassword = "Pass";
				LogFactory.info("Password ... " + wbelTxtPasswordName.getText());
			}
			if (strFlagUserName.equalsIgnoreCase("Pass") && strFlagPassword.equalsIgnoreCase("Pass")) {

				wbelBtnSignin.click();
				LogFactory.info("Clicked on Sign-In button... ");
				return true;
			}
		} catch (Exception e) {
			System.out.println("############################" + e.getMessage());

		}
		return false;

	}

	public static boolean verifyUserLogin() throws Throwable {

		String strFlag = "FAIL";
		String strResult = "User unable to login";
		try {

			if (homePageFactory.wbelUserInfo.isDisplayed()) {
				//LogFactory.info("User logged-in successfully");
				strResult = "User logged-in successfully";
				strFlag = "PASS";

			}
			
			System.out.println("Login & Homepage, Verify User login "+BaseClass.strUserName+"\n"+" User should redirect to homepage");
			
			if (strFlag.equalsIgnoreCase("PASS")) {

				return true;
			}
		} catch (Exception e) {

				
				System.out.println("Login & Homepage, Verify User login"+BaseClass.strUserName+"\n"+" User should redirect to homepage"+ e.getMessage().toString());
		}

		return false;

	}

	public static boolean verifyImpersonatedUser() throws Throwable {

		String strFlag = "FAIL";
		String strResult = "User unable to login";
		try {

			if (Homepage_POF.wbelUserInfo.isDisplayed()) {
				LogFactory.info("User Impersonated Successfully");
				strResult = "User Impersonated Successfully";
				strFlag = "PASS";

			}

			ReportFactory.tableEnd();
			GenericFactory.createHeaderSection("User " + BaseClass.strUserRACFID + " Impersonated");
			ReportFactory.reporterOutput("Login & Homepage", "Verify user impersonate", BaseClass.strUserRACFID,
					"User should redirect to User homepage", strResult, strFlag);

			if (strFlag.equalsIgnoreCase("PASS")) {
				return true;
			}

		} catch (Exception e) {
			ReportFactory.reporterOutput("Login & Homepage", "Verify user login", BaseClass.strUserRACFID,
					"User should redirect to homepage", e.getMessage().toString(), strFlag);
		}
		return false;
	}
	
	
	
	public static void navigateToWCM() throws Throwable
	{
					
	try
	{
			System.out.println("**Inside navigate WCM method**");
		WebElement userelement = null;
		
		String parentHandle = locDriver.getWindowHandle();
		
		
		userelement = ValidationFactory.getElementIfPresent(By.xpath("//div[contains (@class,'user-info') and @id ='js-user-info']"));

		if (userelement != null) {

			userelement.click();
			System.out.println("***Utility Menu Clicked.....");
			
			
			if(ValidationFactory.isElementPresent(wcmLink))
			{
			wcmLink.click();
			System.out.println("***WCM link clicked......");
			}	
			//locDriver.findElement(By.xpath("//a[contains (.,'WCM')]")).click();
//			/System.out.println("***WCM link clicked......");
			
			Thread.sleep(5000);
			
		
				Set<String> handles = locDriver.getWindowHandles();
				//System.out.println("Total opened windows are: "+handles.size());
		
				Iterator<String> it = handles.iterator();
				
				
				while(it.hasNext()) {
					String   parent= it.next();
					String child = it.next();
					
					locDriver.switchTo().window(child);
					
					locDriver.manage().window().maximize();
					
				}
			
		}
	}
	catch(Exception e)
	{
		System.out.println("Link not clicked "+e.getMessage().toString());
	}
	
	
}

	
	
	
	
}
