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
import com.deere.Helpers.ValidationFactory;

/* 
 * Project    : DealerPath
 * Script     : PortalLeftNavigation_POF 
 * Author     : Shrey choudhary
 * Date       : May.15.2018
 * Last Modified On: May.15.2018
 * Modified By : Shrey choudhary
 */
public class PortalLeftNavigation_POF {
	final WebDriver locDriver;

	public PortalLeftNavigation_POF(WebDriver driver) {

		this.locDriver = driver;

	}

	@FindBy(how = How.ID, using = "leftNav")
	static WebElement wbelLeftWindowFrame;

	@FindBy(how = How.XPATH, using = ".//*[@id='leftNav']/li/a[@class='active']")
	public static List<WebElement> ListAllActiveLinks;

	@FindBy(how = How.XPATH, using = ".//*[@id='leftNav']/li/a[@class='inactive']")
	public static List<WebElement> listAllInactiveLinks;

	/**
	 * This method verifies Left navigation window link names are in correct order
	 * 
	 * @author shrey.choudhary
	 * @createdAt 15-05-2018
	 * @param testData
	 * @param TCID
	 * @return
	 * @throws Throwable
	 * @modifiedAt 21-05-2018
	 */
	public static void compareNavigationLinksWithTestData(String strTestData, String strTCID) throws Throwable {
		String strFlag = "Fail";
		List<String> listExpectedData = null;
		List<String> listActualData = null;
		
		String result = "Portal Left Navigation links are not matching with the expected list ";

		try {
			if (ValidationFactory.isElementPresent(wbelLeftWindowFrame)) {
				listExpectedData = GenericFactory.splitString(strTestData, ",");
				
				listActualData = new ArrayList<String>();
				
				for (int i = 0; i < ListAllActiveLinks.size(); i++) {
					String temp = ListAllActiveLinks.get(i).getText();
					listActualData.add(temp);
				}
				if (listExpectedData.equals(listActualData)) {
					strFlag = "Pass";
					result = "Links order names should be same" + listActualData.toString();
				}
				ReportFactory.reporterOutput(strTCID, "Verify Portal Left Navigation Links names ",listExpectedData.toString(),
						"Left Navigation links and their order should match with expected list" , result + "Actual Left Navigation Links : "+ listActualData, strFlag);
			} else {
				ReportFactory.reporterOutput(strTCID, "Verify Portal Left Navigation Links name", strTestData,
						"Verify the order of links name " + listExpectedData.toString(), "Portal Left Navigation window is not present", strFlag);

				/*
				 * if (strFlag.equalsIgnoreCase("FAIL") ) { Assert.assertFalse(true);}
				 */

			}
		} catch (Exception e) {

			ReportFactory.reporterOutput(strTCID, "Verify Portal Left Navigation Links name",
					"Verify Portal Left Navigation Links name", "NA", e.getMessage().toString(), strFlag);
		}
	}
	
	
	
	
	

}
