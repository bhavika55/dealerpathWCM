/* 
 * Project    : DealerPath
 * Script     : Homepage_POF
 * Author     : Neeraja Mantri
 * Date       : May.15.2018
 * Last Modified On:
 * Modified By :
 */

package com.deere.PageFactory;

import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map.Entry;

import org.openqa.selenium.By;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.How;
import org.testng.Assert;
import org.testng.asserts.SoftAssert;

import com.deere.Helpers.BaseClass;
import com.deere.Helpers.DateFactory;
import com.deere.Helpers.GenericFactory;
import com.deere.Helpers.LinkFactory;
import com.deere.Helpers.LogFactory;
import com.deere.Helpers.ReportFactory;
import com.deere.Helpers.ValidationFactory;
import com.deere.Helpers.WaitFactory;

/**
 * @author neeraja.mantri
 *
 */
public class Homepage_POF {

	static WebDriver HomDriver;
	// static SoftAssert softAssert = new SoftAssert();

	public Homepage_POF(WebDriver driver) {
		this.HomDriver = driver;
	}

	@FindBy(how = How.XPATH, using = "//div[@class='user-info' ]")
	static WebElement wbelUserInfo;

	@FindBy(how = How.XPATH, using = "//h1[@class='app-title']")
	static WebElement wbelTitleOfHomePage;

	@FindBy(how = How.ID, using = "leftNav")
	static WebElement wbelLeftWindow;

	@FindBy(how = How.XPATH, using = ".//*[@class='wpthemeFooter']")
	public static WebElement wbelHomePageFooterFramePath;

	@FindBy(how = How.XPATH, using = "//div[@id='js-notifications']")
	static WebElement wbNotificationLink;

	@FindBy(how = How.ID, using = ".//*[@class='list-item']/div")
	static List<WebElement> wbNotificationList;

	@FindBy(how = How.XPATH, using = ".//*[@id='layoutContainers']/div/div[2]/div[3]/div[2]/section/div/div[2]/div")
	static WebElement alertIndex;

	@FindBy(how = How.XPATH, using = ".//*[@id='layoutContainers']/div/div[2]/div[3]/div[3]/section/div/div[2]/div")
	static WebElement announcementIndex;

	@FindBy(how = How.XPATH, using = ".//*[@id='layoutContainers']/div/div[2]/div[3]/div[4]/section/div/div[2]/div")
	static WebElement favoritesIndex;

	@FindBy(how = How.XPATH, using = ".//*[@id='layoutContainers']/div/div[2]/div[3]/div[1]")
	static WebElement searchIndex;

	@FindBy(how = How.XPATH, using = "//div[@class='section-header']/h3")
	static List<WebElement> sectionHeader;

	@FindBy(how = How.ID, using = "//div[@class='value']//input[@type='radio' and @checked='checked']")
	static WebElement wbelSelectedValue;

	@FindBy(how = How.XPATH, using = ".//*[@class='wpthemeFooter']")
	public static WebElement homepageFooterFramePath;

	@FindBy(how = How.XPATH, using = ".//*[@id='dealerPathFooterCopyright']")
	public static WebElement copyrightFooterPath;

	@FindBy(how = How.XPATH, using = ".//*[@id='main-header']/div[1]/div/h1")
	static WebElement homepagepath;

	/**
	 * @param title
	 * @param TCID
	 * @throws Throwable
	 */
	public static void checkUserLogIntoHomepage(String strTitle, String strTCID) throws Throwable {

		String strPageTitle = HomDriver.getTitle();
		String strPageName = "";

		WaitFactory.waitForPageLoaded();

		LogFactory.info(" Page Title " + strPageTitle + " is displayed on the homepage");
		String strFlag = "FAIL";
		String strResult = "Page Title & Header are not showing as expected";

		try {

			if (ValidationFactory.isElementPresent(wbelTitleOfHomePage)) {
				strPageName = wbelTitleOfHomePage.getText();

				if (strPageTitle.equalsIgnoreCase(strTitle) && strPageName.equalsIgnoreCase(strTitle)) {
					strFlag = "PASS";
					strResult = "Page Title  and Page Header are showing as expected";
				}

			}

			ReportFactory.reporterOutput(strTCID, "Verify the Page Title & Header Name on the homepage", strTitle,
					"Page Title & Page Header on Homepage should be " + strPageTitle, strResult, strFlag);

		} catch (Throwable e) {

			ReportFactory.reporterOutput(strTCID, "Verify page title and page header on home page", "NA", "NA",
					e.getMessage().toString(), strFlag);
		}

	}

	/**
	 * @param welcomemsg
	 * @param TCID
	 * @throws Throwable
	 */
	public static void getWelcomeMessageOnHomepage(String strWelcomemsg, String strTCID) throws Throwable {

		String strWelcomeMsg = wbelUserInfo.getText();
		String strWelcome="";
		
		LogFactory.info(" Welcome username :" + strWelcomeMsg);
		String strFlag = "FAIL";
		String strResult = "Welcome message is not displayed";

		try {
			if (ValidationFactory.isElementPresent(wbelUserInfo)) {
				
				 strWelcome = (GenericFactory.splitString(strWelcomeMsg, ",").get(0).toString().trim());

				if (strWelcomeMsg.contains(strWelcome)) {
					strFlag = "PASS";
					strResult = "Welcome message is displayed :" + strWelcome;
				}
			}

			ReportFactory.reporterOutput(strTCID, "Verify welcome message on homepage", strWelcomemsg,
					"Welcome message on homepage should be displayed ", strResult, strFlag);

		} catch (Exception e) {

			ReportFactory.reporterOutput(strTCID, "Verify welcome message on homepage", strWelcomemsg,
					"Welcome message on homepage should be displayed ", e.getMessage().toString(), strFlag);
		}

	}

	/**
	 * @author shrey.choudhary
	 * @createdAt 22-05-2018
	 * @param testData
	 * @param TCID
	 * @throws Throwable
	 * @modifiedAt 22-05-2018
	 */

	public static void getDealerPrincipalRole(String strDealerPrincipleRole, String strTestData, String strTCID)
			throws Throwable {

		String strFlag = "FAIL";
		String strResult = "Dealer is not having dealer principal role";

		try {

			if (strTestData.equals("NA")) {
				ReportFactory.reporterOutput(strTCID, "Verify Dealer Principal role.", strTestData,
						"Verify Dealer Principal role.", "Invalid TestData", strFlag);
			} else {

				List<String> frameList = new ArrayList<String>();
				for (int i = 0; i < PortalLeftNavigation_POF.ListAllActiveLinks.size(); i++) {
					String temp = PortalLeftNavigation_POF.ListAllActiveLinks.get(i).getText();
					frameList.add(temp);
				}
				if (frameList.contains(strTestData) && strDealerPrincipleRole.equalsIgnoreCase("Yes")) {
					strFlag = "Pass";
					strResult = "Dealer has the Dealer Priciple role and also link Dealer Priciple found in the Left Naviagation";
				}

				else if (frameList.contains(strTestData) && strDealerPrincipleRole.equalsIgnoreCase("No")) {
					strFlag = "Fail";
					strResult = "Dealer does not have Dealer Priciple role but link Dealer Priciple found in the Left Naviagation";
				}

				else {
					strFlag = "Fail";
					strResult = "Dealer has the Dealer Priciple role but the link Dealer Priciple not found in the Left Naviagation";
				}

				ReportFactory.reporterOutput(strTCID, "Verify Dealer Principal role.", strTestData,
						"Verify Dealer Principal role.", strResult, strFlag);
			}

		} catch (Throwable e) {

			ReportFactory.reporterOutput(strTCID, "Verify Dealer Principal role.", strTestData,
					"Verify Dealer Principal role.", e.getMessage().toString(), strFlag);
		}
	}

	/**
	 * This method verifies Notification icon is present in user's homepage.
	 * 
	 * @author shrey.choudhary
	 * @createdAt 22-05-2018
	 * @param strNotification
	 * @param TCID
	 * @throws Throwable
	 * @modifiedAt 22-05-2018
	 */
	public static void verifyNotificationIconOnHomePage(String TCID) throws Throwable {
		String flag = "Fail";
		String result = "Notification link is not applicable/available";

		try {
			// if (ValidationFactory.isElementPresent(notificationLink))

			WebElement ele = WaitFactory.waitForElementClickable(wbNotificationLink);
			Thread.sleep(5000);
			if (ValidationFactory.isElementPresent(ele)) {
				result = "Notification link is available";
				flag = "Pass";
			}

			ReportFactory.reporterOutput(TCID, "Verify notification icon on homepage.", "NA",
					"Verify notification link on homepage.", result, flag);
		} catch (NoSuchElementException e) {

			ReportFactory.reporterOutput(TCID, "Verify notification icon on homepage.", "NA",
					"Verify notification link on homepage.", result, flag);
		} catch (Exception e) {
			ReportFactory.reporterOutput(TCID, "Verify notification icon on homepage.", "NA",
					"Verify notification link on homepage.", e.getMessage().toString().substring(0, 125), flag);
		}
	}

	/**
	 * This method compares Notification list with the given input through Excel
	 * 
	 * @author shrey.choudhary
	 * @createdAt 22-05-2018
	 * @param strNotification
	 * @param TCID
	 * @throws Throwable
	 * @modifiedAt 22-05-2018
	 */
	public static void verifyNotificationsList(String strNotification, String TCID) throws Throwable {

		String flag = "Fail";
		try {
			// if (ValidationFactory.isElementPresent(notificationLink))
			if (wbNotificationLink.isDisplayed()) {
				wbNotificationLink.click();
				List<String> ExpectedData = GenericFactory.splitString(strNotification, ",");
				if (!strNotification.equalsIgnoreCase("NA")) {
					List<WebElement> wbNotification = BaseClass.wbDriver
							.findElements(By.xpath(".//*[@class='list-item']/div"));
					List<String> notificationList = new ArrayList<String>();
					for (int i = 0; i < wbNotification.size(); i++) {
						String temp = wbNotification.get(i).getText();
						notificationList.add(temp);
					}
					// System.out.println(notificationList);
					if (notificationList.equals(ExpectedData)) {
						flag = "Pass";
					}
					ReportFactory.reporterOutput(TCID, "Verify notification link names on homepage.",
							"Input data is : " + ExpectedData, "Notification list should match with Input data",
							"Actual data : " + notificationList, flag);
				}

			} else {
				ReportFactory.reporterOutput(TCID, "Verify notification link names on homepage.", "Notification List",
						"Notification list should match with Input data", "Notification List is not present.", flag);
			}
		}

		catch (Exception e) {
			ReportFactory.reporterOutput(TCID, "Verify notification link names on homepage.", "NA",
					"Verify notification link names on homepage.", e.getMessage().toString().substring(0, 125), flag);
		}
	}

	/*
	 * public static void verifyFooterLinksonHomepage(String header, String TCID)
	 * throws Throwable { // String TCID= ""; String flag = "Fail"; String result =
	 * "FooterLinks are not showing as expected on UI."; ArrayList<String>
	 * StrhomepageFooterList = new ArrayList<String>();
	 * 
	 * try {
	 * 
	 * List<WebElement> WeHomepageFooterList =
	 * GenericFactory.getLinksFromFrame(wbelHomePageFooterFramePath); //
	 * List<WebElement> WeHomepageFooterList = //
	 * BaseClass.wbDv.findElements(By.xpath(
	 * "html/body/div[15]/footer/div/div/ul/li/a")); for (int i = 0; i <
	 * WeHomepageFooterList.size(); i++) { String hearlinktxtvalue =
	 * WeHomepageFooterList.get(i).getText().toString().trim();
	 * StrhomepageFooterList.add(hearlinktxtvalue); }
	 * 
	 * // System.out.println(StrhomepageFooterList);
	 * 
	 * // To compare with testdata String expectedtestdata = header; String[]
	 * StrExpectedtestFooterdataList = expectedtestdata.split(",");
	 * 
	 * for (int i = 0; i < StrhomepageFooterList.size(); i++) { String expectedvalue
	 * = StrExpectedtestFooterdataList[i];
	 * 
	 * if (expectedvalue.contains("Switch Site")) { flag = "Pass";
	 * LogFactory.info(expectedvalue + " is present on Footer link"); } else { if
	 * (StrhomepageFooterList.contains(expectedvalue)) { flag = "Pass";
	 * LogFactory.info(expectedvalue + " is present on Footer link"); } else { flag
	 * = "Fail"; LogFactory.info(expectedvalue + " is not present on Footer link");
	 * break; } }
	 * 
	 * // To check system date and match the date with the 'Copyright © 2017 Deere &
	 * // Company. All Rights Reserved' link
	 * 
	 * // GenericFactory.isValidDateFormat(dateFormat, listOfDates) String
	 * copyrighttxt = BaseClass.wbDriver
	 * .findElement(By.xpath("html/body/div[15]/footer/div/div[2]/ul/li[3]/a")).
	 * getText().toString() .trim();
	 * 
	 * List<String> copyrightYr = GenericFactory.splitString(copyrighttxt, " ");
	 * String copyrightYrValue = copyrightYr.get(2);
	 * System.out.println(copyrightYrValue);
	 * 
	 * // getting system DateTime Date date = new Date(); String systemYr =
	 * DateFactory.captureYear(date);
	 * 
	 * // getting UI 'copyright' link date if (copyrightYrValue.equals(systemYr)) {
	 * LogFactory.info("Copyright date is correct as FY yr.");
	 * 
	 * } else { LogFactory.info("Copyright date is not correct as FY yr."); }
	 * 
	 * }
	 * 
	 * result = "FooterLinks is showing as per on the UI ";
	 * ReportFactory.reporterOutput(TCID, "verify FooterLinks on Homepage", "NA",
	 * "FooterLinks should  show as per on the UI ", result, flag);
	 * 
	 * 
	 * } catch (Exception e) { // LogFactory.error("e"); String er =
	 * e.getMessage().substring(0, 35).toString().trim();
	 * 
	 * ReportFactory.reporterOutput(TCID, "verify FooterLinks on Homepage", "NA",
	 * "NA", er, flag); }
	 * 
	 * // *************************************************** }
	 */

	/*
	 * // To verify homepage footer Links public static void
	 * verifyFooterLinksonHomepage(String header, String TCID) throws Throwable {
	 * 
	 * String flag = "Fail"; String result =
	 * "Footer links are not showing as expected"; ArrayList<String>
	 * StrhomepageFooterList = new ArrayList<String>();
	 * 
	 * try {
	 * 
	 * List<WebElement> WeHomepageFooterList =
	 * GenericFactory.getLinksFromFrame(homepageFooterFramePath); //
	 * List<WebElement> WeHomepageFooterList = //
	 * BaseClass.wbDv.findElements(By.xpath(
	 * "html/body/div[15]/footer/div/div/ul/li/a")); for (int i = 0; i <
	 * WeHomepageFooterList.size(); i++) { String hearlinktxtvalue =
	 * WeHomepageFooterList.get(i).getText().toString().trim();
	 * StrhomepageFooterList.add(hearlinktxtvalue); }
	 * 
	 * // System.out.println("Str homepage FooterList= " +StrhomepageFooterList);
	 * 
	 * // To compare with testdata String expectedtestdata = header; String[]
	 * StrExpectedtestFooterdataList = expectedtestdata.split(",");
	 * 
	 * for (int i = 0; i < StrhomepageFooterList.size(); i++) { String expectedvalue
	 * = StrExpectedtestFooterdataList[i];
	 * 
	 * if (expectedvalue.contains("Switch Site")) { flag = "Pass";
	 * LogFactory.info(expectedvalue + " is present on Footer link"); } else if
	 * (StrhomepageFooterList.contains(expectedvalue)) { flag = "Pass"; result =
	 * "Footer links are showing as expected"; LogFactory.info(expectedvalue +
	 * " is present on Footer link");
	 * 
	 * } else { flag = "Fail"; break; } }
	 * 
	 * 
	 * ReportFactory.reporterOutput(TCID,
	 * "Verify the sequence & presence of Footer links on the homepage ", header,
	 * "Footer links should be present and show in sequence ", result, flag);
	 * 
	 * flag = "Fail";
	 * 
	 * // getting system DateTime String copyText="";
	 * 
	 * for(int i=0; i<StrhomepageFooterList.size(); i++) { copyText =
	 * StrhomepageFooterList.get(i).toString().trim();
	 * if(copyText.contains("Copyright")) { List<String> copyrightYr =
	 * GenericFactory.splitString(copyText, " "); String copyrightYrValue =
	 * copyrightYr.get(2); Date date = new Date(); String systemYr =
	 * DateFactory.captureYear(date);
	 * 
	 * // getting UI 'copyright' link date if (copyrightYrValue.equals(systemYr)) {
	 * flag = "Pass"; result = "Copyright is displayed with current year ";
	 * LogFactory.info("Copyright year is correct as FY yr.");
	 * 
	 * } else {
	 * 
	 * flag = "Fail"; result = "Copyright is displayed with incorrect year";
	 * LogFactory.info("Copyright year is not correct as FY yr.");
	 * 
	 * } } }
	 * 
	 * ReportFactory.reporterOutput("TC26_Homepage",
	 * "Verify copyright link is displayed as per current year", header,
	 * "Copyright link should be displayed with current year", result, flag);
	 * 
	 * 
	 * } catch (Exception e) { // LogFactory.error("e"); String er =
	 * e.getMessage().substring(0, 35).toString().trim();
	 * 
	 * ReportFactory.reporterOutput(TCID,
	 * "Verify the sequence & presence of Footer links on the homepage", "NA",
	 * "FooterLinks should show as per on the UI.", er, flag);
	 * 
	 * } }
	 * 
	 */

	// ***************************************************** E-O-M
	// *****************************************************************************

	/*
	 * @author sweta.ranjan
	 * 
	 * @createdAt 04-05-2018
	 * 
	 * @throws IOException
	 * 
	 * @throws Exception
	 * 
	 * @modifyBy sweta.ranjan
	 * 
	 * @modifyAt 12-07-2018
	 */

	// To verify homepage footer Links
	public static void verifyFooterLinksonHomepage(String header, String TCID) throws Throwable {
		// String TCID = "Footer Links";
		String flag = "Fail";
		String result = "Footer links are not showing as expected";
		ArrayList<String> StrhomepageFooterList = new ArrayList<String>();

		try {

			List<WebElement> WeHomepageFooterList = GenericFactory.getLinksFromFrame(homepageFooterFramePath);
			// List<WebElement> WeHomepageFooterList =
			// BaseClass.wbDv.findElements(By.xpath("html/body/div[15]/footer/div/div/ul/li/a"));
			for (int i = 0; i < WeHomepageFooterList.size(); i++) {
				String hearlinktxtvalue = WeHomepageFooterList.get(i).getText().toString().trim();
				StrhomepageFooterList.add(hearlinktxtvalue);

			}

			System.out.println("Str homepage FooterList= " + StrhomepageFooterList);

			// To compare with testdata
			String expectedtestdata = header;
			String[] StrExpectedtestFooterdataList = expectedtestdata.split(",");

			for (int i = 0; i < StrhomepageFooterList.size(); i++) {
				String expectedvalue = StrExpectedtestFooterdataList[i];
				System.out.println(StrhomepageFooterList.get(i).trim() + "--" + expectedvalue.trim());

				if (StrhomepageFooterList.get(i).trim().contains(expectedvalue.trim())) {
					flag = "Pass";
					result = "Footer links are showing as expected";
					LogFactory.info(expectedvalue + " is present on Footer link");
				} else {
					flag = "Fail";
					result = "FooterLinks is not showing as expected on the UI ";
					LogFactory.info(expectedvalue + " is not present on Footer link ");

					break;
				}
			}

			ReportFactory.reporterOutput(TCID, "Verify the sequence & presence of Footer links on the homepage ",
					header, "Footer links should be present and show in sequence ",
					result + " Actual Footer Links : " + StrhomepageFooterList, flag);

			flag = "Fail";

		} catch (Exception e) {
			// LogFactory.error("e");
			String er = e.getMessage().substring(0, 35).toString().trim();

			ReportFactory.reporterOutput(TCID, "Verify the sequence & presence of Footer links on the homepage", "NA",
					"FooterLinks should show as per on the UI.", er, flag);

		}
	}

	// ***************************************************** E-O-M
	// *****************************************************************************

	// ***************************************************** E-O-M
	// *****************************************************************************

	/*
	 * @author sweta.ranjan
	 * 
	 * @createdAt 04-05-2018
	 * 
	 * @throws IOException
	 * 
	 * @throws Exception
	 * 
	 * @modifyBy sweta.ranjan
	 * 
	 * @modifyAt 12-07-2018
	 */

	// To verify homepage footer Links
	public static void verifyCopyrightOfFooterLinksonHomepage(String TCID) throws Throwable {
		// String TCID = "Copyright Footer Links";
		String flag = "Fail";
		String result = "Copyright is displayed with incorrect year";

		// Store all footer links in an arraylist
		ArrayList<String> StrhomepageFooterList = new ArrayList<String>();

		// getting system DateTime
		String actualCopyText = "";
		// String expectedCopyText= "Copyright";
		String copyrirightText = copyrightFooterPath.getText().toString().trim();
		try {
			List<String> copyrightYr = GenericFactory.splitString(copyrirightText, " ");
			String copyrightYrValue = copyrightYr.get(2);
			Date date = new Date();
			String systemYr = DateFactory.captureYear(date);

			// getting UI 'copyright' link date
			if (copyrightYrValue.equals(systemYr)) {
				flag = "Pass";
				result = "Copyright is displayed with current year ";
				LogFactory.info("Copyright Date is correct as FY yr.");

			} else {
				flag = "Fail";
				result = "Copyright is displayed with incorrect year";
				LogFactory.info("Copyright Date is not correct as FY yr.");
			}

			ReportFactory.reporterOutput(TCID, "Verify copyright link is displayed as per current year", "NA",
					"Copyright link should be displayed with current year", result, flag);

		} catch (Exception e) {
			// LogFactory.error("e");
			String er = e.getMessage().substring(0, 35).toString().trim();

			ReportFactory.reporterOutput(TCID, "verify copyright footerLinks on homepage", "NA",
					"Copyright footerLinks should show as per on the UI.", er, flag);

		}
	}

	// ***************************************************** E-O-M
	// *****************************************************************************

	public static void checkorderofportletsdisplayed(String TCID) throws Throwable {
		String flag = "Fail";
		String resultPortletOrder = "";
		try {
			// List adding portlet frames
			ArrayList<WebElement> listPortletframes = new ArrayList<WebElement>();
			listPortletframes.add(searchIndex);
			listPortletframes.add(alertIndex);
			listPortletframes.add(announcementIndex);
			listPortletframes.add(favoritesIndex);

			int intListPortletFrameSize = listPortletframes.size();

			HashMap<Integer, String> headerTxtHomePage = new HashMap<Integer, String>();
			int intListPageSize = sectionHeader.size();

			for (int i = 0; i < sectionHeader.size(); i++) {
				headerTxtHomePage.put(i, sectionHeader.get(i).getText());

				if (sectionHeader.get(i).getText().contains("My DealerPath Announcements")) {

					String strAnnouncementHeaderTxt = sectionHeader.get(i).getText();
					String strAnnouncementTxt = strAnnouncementHeaderTxt.substring(0,
							strAnnouncementHeaderTxt.lastIndexOf('(') - 1);

					headerTxtHomePage.put(i, strAnnouncementTxt);
				}

				// System.out.println("List of elements :" + sectionHeader.get(i).getText());
			}
			HashMap<Integer, String> headerTxtExcelValues = new HashMap<Integer, String>();
			headerTxtExcelValues.put(0, "DealerPath Alerts");
			headerTxtExcelValues.put(1, "My DealerPath Announcements");
			headerTxtExcelValues.put(2, "My DealerPath Favorites");

			int headerValue = headerTxtHomePage.size();

			switch (headerValue) {
			case 3:
				resultPortletOrder = "Portlets are not displayed as expected";
				for (int i = 0; i < headerTxtExcelValues.size(); i++) {
					if (headerTxtExcelValues.get(i).equals(headerTxtHomePage.get(i))) {
						flag = "Pass";
						resultPortletOrder = "Alerts,Announcements & Favourites portlet order are displayed as expected";
						LogFactory.info(
								headerTxtExcelValues.get(i) + " " + "are equal" + " " + headerTxtExcelValues.get(i));
					}

					else {
						LogFactory.info(headerTxtExcelValues.get(i) + " " + "are not equal" + " "
								+ headerTxtExcelValues.get(i));
						flag = "Fail";
						break;
					}

				}
				LogFactory.info("All the portlets are present in order");
				ReportFactory.reporterOutput(TCID,
						"Verify Portlets displayed in order :Alert portlet first,announcement portlet second and favourites portlet last.",
						"NA",
						"Portlets should be displayed in order :Alert portlet in first,announcement portlet in second and favourites portlet in last",
						resultPortletOrder, flag);
				break;
			case 2:

				resultPortletOrder = "Portlets are not displayed as expected";

				for (int i = 0; i < headerTxtHomePage.size(); i++) {

					String strHeaderTxt = headerTxtHomePage.get(i);
					if (strHeaderTxt.equals("DealerPath Alerts")) {
						flag = "Pass";
						resultPortletOrder = "Alert portlet is displayed first and above announcements & favourites portlet";
						LogFactory.info("ALerts portlet is displayed first ");
						ReportFactory.reporterOutput(TCID,
								"Verify Portlets displayed in order :Alert portlet first,announcement portlet second and favourites portlet last.",
								"NA",
								"Portlets should be displayed in order :Alert portlet in first,announcement portlet in second and favourites portlet in last",
								resultPortletOrder, flag);
						// break;
					} else if (strHeaderTxt.equals("My DealerPath Announcements")) {
						flag = "Pass";
						resultPortletOrder = "Announcement portlet is displayed as expected and above favourites portlet or below alerts portlet";
						LogFactory.info("Announcements portlet is displayed");
						ReportFactory.reporterOutput(TCID,
								"Verify Portlets displayed in order :Alert portlet first,announcement portlet second and favourites portlet last.",
								"NA",
								"Portlets should be displayed in order :Alert portlet in first,announcement portlet in second and favourites portlet in last",
								resultPortletOrder, flag);
						break;
					} else if (strHeaderTxt.equals("My DealerPath Favourites")) {
						flag = "Pass";
						LogFactory.info("Favourites portlet is displayed");
						resultPortletOrder = "Favorites portlet is displayed at last";
						ReportFactory.reporterOutput(TCID,
								"Verify Portlets displayed in order :Alert portlet first,announcement portlet second and favourites portlet last.",
								"NA",
								"Portlets should be displayed in order :Alert portlet in first,announcement portlet in second and favourites portlet in last",
								resultPortletOrder, flag);
						break;
					} else {
						flag = "Fail";
						ReportFactory.reporterOutput(TCID,
								"Verify Portlets displayed in order :Alert portlet first,announcement portlet second and favourites portlet last.",
								"NA",
								"Portlets should be displayed in order :Alert portlet in first,announcement portlet in second and favourites portlet in last",
								resultPortletOrder, flag);
						break;
					}

				}

				break;
			case 1:

				for (int i = 0; i < headerTxtHomePage.size(); i++) {
					String header = headerTxtHomePage.get(i);
					LogFactory.info("Only" + header + "portlet is present");
					resultPortletOrder = "Only" + header + "portlet is displayed as expected.";
					flag = "Pass";
				}
				ReportFactory.reporterOutput(TCID, "Verify order of portlets displayed", "NA",
						"Portlets should be displayed in order :Alert portlet in first,announcement portlet in second and favourites portlet in last",
						resultPortletOrder, flag);

				break;
			default:

				String resultOrder = "No portlets are displayed";
				LogFactory.info("No portlets are displayed");
				flag = "Fail";
				ReportFactory.reporterOutput(TCID, "Verify order of portlets displayed", "NA",
						"Portlets should be displayed in order :Alert portlet in first,announcement portlet in second and favourites portlet in last",
						resultOrder, flag);
			}

		} catch (Exception e) {
			ReportFactory.reporterOutput(TCID, "Verify order of portlets displayed", "NA",
					"Portlets should be displayed in order :Alert portlet in first,announcement portlet in second and favourites portlet in last",
					e.getMessage().toString().substring(0, 125), flag);
		}
	}

	// Verify selected theme color on My preference is as per expected value
	// ********************************************************************************************************************************************

	public static void verifyUserSelectdTheme(String strExpectedValue, String strTCID) throws Throwable {

		String strResult = "Selected theme color on My Preferences is not as Expected";
		String strFlag = "Fail";

		GenericFactory.utilityMenuMyPreferenceClick();

		WebElement wbelSelectedValue = HomDriver
				.findElement(By.xpath("//div[@class='group-value radio-value']//div[1]//label[1]//div"));

		
		String strSelectedValue = wbelSelectedValue.getText().trim();
		
		
		if (strSelectedValue.equalsIgnoreCase(strExpectedValue.trim())) {
			strFlag = "Pass";
			strResult = "Selected theme color on My Preferences is as Expected";
		}

		ReportFactory.reporterOutput(strTCID, "Verify text of selected theme color", strExpectedValue,
				"Selected theme color text should be as per Input value", strResult, strFlag);

		HomDriver.findElement(By.xpath("//div[@id='my-preferences']//button[@class='close']//span")).click();

	}

}

// *********************************************************************************************************************************************
