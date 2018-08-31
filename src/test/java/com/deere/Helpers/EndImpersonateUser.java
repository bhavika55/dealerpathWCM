package com.deere.Helpers;

import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.testng.Assert;
import org.testng.annotations.AfterClass;

/**
 * This method is used to end impersonate the dealer
 * 
 * @author shrishail.baddi
 * @param strRACFID
 * @createdAt 05-06-2018
 * @throws IOException
 * 
 * @modifyBy shrey.choudhary
 * @modifyAt 07-06-2018
 */
public class EndImpersonateUser {
	public static void endImpersonateOrLogoutUser() {
		WebElement userelement = null;
		try {
			if (BaseClass.errorUserFound == false) {
				userelement = ValidationFactory
						.getElementIfPresent(By.xpath("//div[contains (@class,'user-info') and @id ='js-user-info']"));
				if (userelement != null) {
					userelement.click();
					WebElement btnelement = ValidationFactory.getElementIfPresent(
							By.xpath("//button[@class='btn' and @onclick='endImpersonate(this)']"));
					if (btnelement != null) {
						btnelement.click();
					} else {
						ValidationFactory
								.getElementIfPresent(
										By.xpath("//button[@class='btn' and @onclick='logoutDealerPath(this)']"))
								.click();
					}
				}
			} else {
				userelement = ValidationFactory
						.getElementIfPresent(By.xpath("//div[contains (@class,'user-info') and @id ='js-user-info']"));
				if (userelement != null) {
					userelement.click();
					ValidationFactory.getElementIfPresent(
							By.xpath("//button[@class='btn' and @onclick='logoutDealerPath(this)']")).click();
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
			LogFactory.info("Unable to logout or impersonate the user");
			//Assert.fail("Unable to logout or impersonate the user");
		}
	}


@AfterClass
public void getReportFooter() throws InterruptedException {
//	LogFactory.endTestCase("Home Page Testcases");
	ReportFactory.tableEnd();

}

}