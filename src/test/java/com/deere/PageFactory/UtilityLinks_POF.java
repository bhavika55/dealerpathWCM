package com.deere.PageFactory;

import java.util.ArrayList;
import java.util.List;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.How;
import org.testng.Assert;

import com.deere.Helpers.GenericFactory;
import com.deere.Helpers.ReportFactory;

/* 
 * Project    : DealerPath
 * Script     : UtilityLink_POF 
 * Author     : Shrey choudhary
 * Date       : May.15.2018
 * Last Modified On: May.15.2018
 * Modified By : Shrey choudhary
 */
public class UtilityLinks_POF {
	final WebDriver locDriver;

	public UtilityLinks_POF(WebDriver driver) {

		this.locDriver = driver;

	}

	@FindBy(how = How.ID, using = "js-user-info")
	static WebElement wbelClickUtilityLinks;

	@FindBy(how = How.ID, using = "js-user-info-pop")
	static WebElement wbelAllUtilityLinks;

	@FindBy(how = How.XPATH, using = "//div[@id='js-user-info-pop']/div[2]/button")
	static WebElement wbelUtilityLinkButton;

	/**
	 * This method verify utility link button name
	 * 
	 * @author shrey.choudhary
	 * @param testData
	 * @param TCID
	 * @throws Throwable
	 */
	public static void compareUtilityLinksWithTestData(String strTestData, String strTCID) throws Throwable {
		String strFlag = "Fail";
		try {
			Thread.sleep(5000);
			wbelClickUtilityLinks.click();
			List<WebElement> listExpectedUtilityElements = GenericFactory.getLinksFromFrame(wbelAllUtilityLinks);
			List<String> listExpectedUtilityLinks = new ArrayList<String>();
			for (int i = 0; i < listExpectedUtilityElements.size(); i++) {
				String temp = listExpectedUtilityElements.get(i).getText();
				listExpectedUtilityLinks.add(temp);
			}
			List<String> ExpectedData = GenericFactory.splitString(strTestData, ",");
			if (listExpectedUtilityLinks.equals(ExpectedData)) {
				strFlag = "Pass";
			}

			ReportFactory.reporterOutput(strTCID, "Verify utility links and their order on the homepage", strTestData,
					"Verify the order of links on utility links on the home page",
					"Order of utility links should be same as " + listExpectedUtilityLinks.toString(), strFlag);

			/*
			 * if (strFlag.equalsIgnoreCase("FAIL") ) { Assert.assertFalse(true);}
			 */

		} catch (Exception e) {

			ReportFactory.reporterOutput(strTCID, "Verify utility links and their order on the homepage",
					"Verify links on the utility menu on homepage", "NA", e.getMessage().toString(), strFlag);
		}

	}

	/**
	 * This method verify utility link button name
	 * 
	 * @author shrey.choudhary
	 * @param testData
	 * @param TCID
	 * @throws Throwable
	 */
	public static void compareUtilityButtonWithTestData(String strTestData, String strTCID) throws Throwable {
		String strFlag = "Fail";
		try {
			String strUtilityButton = wbelUtilityLinkButton.getText();
			String strExpectedButtonName = strTestData;
			if (strUtilityButton.equals(strExpectedButtonName)) {
				strFlag = "Pass";
			}
			ReportFactory.reporterOutput(strTCID,
					"Verify the Signout/EndImpersonate button on the Utility Links Menu on Home Page", strTestData,
					"Signout/EndImpersonate button should be available", strUtilityButton.toString(), strFlag);

/*			if (strFlag.equalsIgnoreCase("FAIL")) {
				Assert.assertFalse(true);
			}*/

		} catch (Exception e) {

			ReportFactory.reporterOutput(strTCID, "Verify Signout/EndImpersonate button on utility link menu", "NA",
					strTestData, e.getMessage().toString(), strFlag);
		}
	}
}
