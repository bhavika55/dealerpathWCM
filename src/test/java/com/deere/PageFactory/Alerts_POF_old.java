/* 
 * Project    : DealerPath
 * Script     : Alerts_POF
 * Author     : Neeraja Mantri
 * Date       : May.15.2018
 * Last Modified On:
 * Modified By :
 */

package com.deere.PageFactory;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

import javax.xml.xpath.XPath;

import org.apache.commons.logging.Log;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.How;
import org.testng.Assert;

import com.deere.Helpers.LogFactory;
import com.deere.Helpers.ReportFactory;
import com.deere.Helpers.ValidationFactory;

public class Alerts_POF_old {

	final WebDriver alrtDriver;

	public Alerts_POF_old(WebDriver driver) {
		this.alrtDriver = driver;

	}

	@FindBy(how = How.XPATH, using = "//div[@class='section warning']")
	static WebElement wbelAlertPortlet;

	@FindBy(how = How.XPATH, using = "//div[@class='section warning']//div[@class='section-header']")
	static WebElement wbelPortletHeader;

	/*
	 * @FindBy (how=How.XPATH, using
	 * ="//h3[@class='section-title'and contains(text(),' DealerPath Alerts')]")
	 * static WebElement alertheader;
	 */

	@FindBy(how = How.XPATH, using = "//span[@class='icon warning']")
	static WebElement wbelWarningImg;

	@FindBy(how = How.XPATH, using = ".//*[@id='layoutContainers']/div/div[2]/div[3]/div[2]/section/div/div[2]")
	public static WebElement wbelAlertFramePath;
	
	

	 

	/**
	 * @throws Throwable
	 */
	public static void isAlertPortletPresent(String TCID) throws Throwable {

		String strFlag = "FAIL";
		String strResult = "Alerts portlet is not present on the homepage";
		try {
			if (ValidationFactory.isElementPresent(wbelAlertPortlet)) {
				strFlag = "PASS";
				strResult = "Alerts portlet is present on the homepage";
			}
			ReportFactory.reporterOutput(TCID, "Verify presence of Alert portlet on Home Page", "NA",
					"Alert portlet should be present on homepage", strResult, strFlag);

			if (strFlag.equalsIgnoreCase("FAIL")) {
				Assert.assertFalse(true);
			}

		} catch (Throwable e) {

			ReportFactory.reporterOutput(TCID, "Verify presence of Alert portlet on Home Page", "NA",
					"Alert portlet should be present on homepage", e.getMessage().toString(), strFlag);
		}

	}

	/**
	 * @param actualheadertxt
	 * @param TCID
	 * @throws Throwable
	 */
	public static void getAlertPortletHeaderAndText(String strActualHeaderTxt, String strTCID) throws Throwable {

		String strFlag = "FAIL";
		String strResult = "Alert header text is not present";
		try {
			if (ValidationFactory.isElementPresent(wbelPortletHeader)) {

				String strAlertHeaderTxt = wbelPortletHeader.getText();
				if (strAlertHeaderTxt.equals(strActualHeaderTxt)) {
					strFlag = "PASS";
					LogFactory.info("Alert portlet contains header with text");
					strResult = "Alert header text is present :" + strActualHeaderTxt;
				}

			}
			ReportFactory.reporterOutput(strTCID, "Verify Alert Header text on homepage", "DealerPath Alerts",
					"Alert header text should be present on homepage", strResult, strFlag);

			if (strFlag.equalsIgnoreCase("FAIL")) {
				Assert.assertFalse(true);
			}

		} catch (Throwable e) {

			ReportFactory.reporterOutput(strTCID, "Verify Alert Header text  on homepage", "DealerPath Alerts",
					"Alert header text should be present on homepage", e.getMessage().toString(), strFlag);

		}

	}

	/**
	 * @throws Throwable
	 */
	public static void checkForWarningSignPresentInAlertHeader(String TCID) throws Throwable

	{
		String strFlag = "FAIL";
		String strResult = "Alert warning sign icon is not displayed";
		try {
			if (ValidationFactory.isElementPresent(wbelWarningImg)) {
				strFlag = "PASS";
				LogFactory.info("Alert warning sign icon is present");
				strResult = "Alert sign is displayed";

			}
			ReportFactory.reporterOutput(TCID, "Verify alert warning sign icon on homepage", "NA",
					"Alert warning sign icon should be present on homepage", strResult, strFlag);

			if (strFlag.equalsIgnoreCase("FAIL")) {
				Assert.assertFalse(true);
			}

		} catch (Throwable e) {

			ReportFactory.reporterOutput(TCID, "Verify alert warning sign icon on homepage", "NA",
					"Alert warning sign icon should be present on homepage", e.getMessage().toString(), strFlag);
		}
	}

	/**
	 * @param alertheadertxtprefrdlang
	 * @param TCID
	 * @throws Throwable
	 */
	public static void getAlertHeaderTxtInPreferredLanguage(String strAlertHeaderTxtPrefrdLang, String strTCID)
			throws Throwable {

		String strAlrtHeaderTxtOnHomePage = null;
		String strFlag = "FAIL";
		String strResult = " Alert headert text is not present in the user preferred language";
		try {
			if (ValidationFactory.isElementPresent(wbelPortletHeader)) {

				LogFactory.info("Alerts portlet header text is present in the user preferred language ");

				strAlrtHeaderTxtOnHomePage = wbelPortletHeader.getText();
				if (strAlrtHeaderTxtOnHomePage.equals(strAlertHeaderTxtPrefrdLang)) {
					strFlag = "PASS";
					strResult = " Alert header text is present in the user preferred language";
				}
				ReportFactory.reporterOutput(strTCID, "Verify alert header text present in the user preferred language",
						strAlrtHeaderTxtOnHomePage,
						"Alert header text should be present in the user preferred language", strResult, strFlag);

				/*
				 * if (strFlag.equalsIgnoreCase("FAIL")) { Assert.assertFalse(true); }
				 */

			}

		} catch (Throwable e) {

			ReportFactory.reporterOutput(strTCID, "Verify alert header text in the user preferred language",
					strAlrtHeaderTxtOnHomePage, "Alert header text should be displayed in preferred language",
					e.getMessage().toString(), strFlag);
		}

	}



}
