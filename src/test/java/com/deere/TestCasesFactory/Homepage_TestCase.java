package com.deere.TestCasesFactory;

import java.util.ArrayList;
import java.util.List;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.PageFactory;
import org.testng.Assert;
import org.testng.annotations.AfterClass;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Parameters;
import org.testng.annotations.Test;

import com.deere.Helpers.BaseClass;
import com.deere.Helpers.BrowserFactory;
import com.deere.Helpers.ExcelFactory;
import com.deere.Helpers.GenericFactory;
import com.deere.Helpers.LogFactory;
import com.deere.Helpers.ReportFactory;
import com.deere.Helpers.ValidationFactory;
import com.deere.Helpers.WaitFactory;
import com.deere.PageFactory.Alerts_POF;
import com.deere.PageFactory.Announcements_POF;
import com.deere.PageFactory.Homepage_POF;
import com.deere.PageFactory.Login_Page_POF;
import com.deere.PageFactory.PortalLeftNavigation_POF;
import com.deere.PageFactory.UtilityLinks_POF;

@Test(groups = { "Homepage" })
public class Homepage_TestCase extends BaseClass {

	WebDriver driver;
	static String strExpectedValue;
	static String strTCID;

	@BeforeClass
	public void getReportHeader() throws InterruptedException {
		ReportFactory.tableEnd();
		GenericFactory.createHeaderSection("Home_Page");
	}

	@Test(priority = 1)
	/* (dependsOnGroups={"Login"}) */

	public void verifyHomepageTitleIsDisplayed() throws Throwable {
		
		strTCID = "TC02_Homepage";
		strExpectedValue = getExcelDataByTestCaseID(strTCID);
		

		if (!strExpectedValue.equalsIgnoreCase("None")) {
			LogFactory.beginTestCase("Verify title on the Homepage is displayed");
			homePageFactory.checkUserLogIntoHomepage(strExpectedValue, strTCID);
		}

	}

	@Test(priority = 2)
	public void verifyWelcomeMessageOnHomepage() throws Throwable {

		strTCID = "TC03_Homepage";
		strExpectedValue = getExcelDataByTestCaseID(strTCID);

		if (!strExpectedValue.equalsIgnoreCase("None")) {
			LogFactory.beginTestCase("Verify welcome message on homepage is showing");
			homePageFactory.getWelcomeMessageOnHomepage(strExpectedValue, strTCID);
		}

	}

	@Test(priority = 3)
	public static void checkVerifyTheme() throws Throwable {
		strTCID = "TC04_Homepage";
		strExpectedValue = getExcelDataByTestCaseID(strTCID);
		LogFactory.beginTestCase("Verify selected theme color label on My Preferences");
		homePageFactory.verifyUserSelectdTheme(strExpectedValue, strTCID);
	}

	@SuppressWarnings("static-access")
	@Test(priority = 4)
	public static void verifyAlertsAreShowingOnHomepage() throws Throwable {

		strTCID = "TC05_Homepage";
		LogFactory.beginTestCase("Verify presence of Alert portlet on Home Page");
		alertPageFactory.isAlertPortletPresent(strTCID);
	}

	@Test(priority = 5)
	public static void verifyALertHeaderWithTxtIsPresentOnHomepage() throws Throwable {

		strTCID = "TC06_Homepage";
		strExpectedValue = getExcelDataByTestCaseID(strTCID);

		if (!strExpectedValue.equalsIgnoreCase("None")) {
			LogFactory.beginTestCase("Verify alerts header text on homepage");
			alertPageFactory.getAlertPortletHeaderAndText(strExpectedValue, strTCID);
		}
	}

	@Test(priority = 6)
	public static void verifyWarningSignPresentWithAlertHeaderOnHomepage() throws Throwable {

		strTCID = "TC07_Homepage";
		LogFactory.beginTestCase("Verify alerts warningsign in alertheader on homepage");
		alertPageFactory.checkForWarningSignPresentInAlertHeader(strTCID);
	}

	@Test(priority = 7)
	public static void verifyHeaderTxtIsInTheUserPrefferedLanguage() throws Throwable {

		strTCID = "TC08_Homepage";
		strExpectedValue = getExcelDataByTestCaseID(strTCID);
		if (!strExpectedValue.equalsIgnoreCase("None")) {
			LogFactory.beginTestCase("Verify alerts headertext is in the user preferred language on Home Page");
			alertPageFactory.getAlertHeaderTxtInPreferredLanguage(strExpectedValue, strTCID);
		}

	}

	@Test(priority = 8)
	public void verifyOrderOfPortlets() throws Throwable {
		strTCID = "TC09_Homepage";
		LogFactory.beginTestCase(strTCID + "Verify Alerts portlet present above announcements portlet");
		homePageFactory.checkorderofportletsdisplayed(strTCID);

	}

	@Test(priority = 9)
	public static void verifyAnnouncementOnTheHomePage() throws Throwable {
		LogFactory.beginTestCase(" Verify presence of Announcements portlet on Home Page");
		strTCID = "TC10_Homepage";
		announcementFactory.verifyAnnouncementTableOnHomePage(strTCID);
	}

	@Test(priority = 10)
	public static void verifyAnnouncementContentIsPresent() throws Throwable {
		strTCID = "TC11_Homepage";
		LogFactory.beginTestCase("Verify Announcement content/body is showing");
		announcementFactory.verifyAnnouncementContentIsPrsent(strTCID);
	}

	@Test(priority = 11)
	public static void verifyAnnouncementHeaderTextIsInUserPrefferedLang() throws Throwable {
		strTCID = "TC12_Homepage";
		strExpectedValue = BaseClass.getExcelDataByTestCaseID(strTCID);
		if (!strExpectedValue.equalsIgnoreCase("None")) {
			LogFactory.beginTestCase("Verify announcement header text in the user preferred language");
		announcementFactory.verifyAnnouncementHeaderTextPrefferedLang(strExpectedValue, strTCID);
		}
	}

	/**
	 * This method verify utility links name should shown in same order as given in
	 * test data sheet
	 * 
	 * @author shrey.choudhary
	 * @throws Throwable
	 */
	@Test(priority = 12)
	public void verifyUtilityMenuLinksOnHomePage() throws Throwable {
		strTCID = "TC13_Homepage";
		strExpectedValue = BaseClass.getExcelDataByTestCaseID(strTCID);

		if (!strExpectedValue.equalsIgnoreCase("None")) {
			LogFactory.beginTestCase("Verify utility links menu and their order on the home page");
			utilityLinksFactory.compareUtilityLinksWithTestData(strExpectedValue, strTCID);
		}
	}

	/**
	 * /**This method verify utility link button name
	 * 
	 * @author shrey.choudhary
	 * @throws Throwable
	 */

	@Test(priority = 13)
	public void verifyUtilityButton() throws Throwable {
		strTCID = "TC14_Homepage";
		strExpectedValue = BaseClass.getExcelDataByTestCaseID(strTCID);

		if (!strExpectedValue.equalsIgnoreCase("None")) {
			LogFactory.beginTestCase("Verify Utility Links menu  button Sign-Out/EndImpersonate");
			utilityLinksFactory.compareUtilityButtonWithTestData(strExpectedValue, strTCID);
		}
	}

	/**
	 * Script : To Test Favourites on Home Page Author : Archana Gaikwad Date :
	 * April.15.2018
	 * 
	 **/
	@SuppressWarnings("static-access")
	@Test(priority = 14)
	public void verifyFavouriteHeaderPresentOnHomePage() throws Throwable {
		strTCID = "TC15_Homepage";
		strExpectedValue = BaseClass.getExcelDataByTestCaseID(strTCID);
		if (!strExpectedValue.equalsIgnoreCase("None")) {
			LogFactory.beginTestCase("Verify favourites header is showing on the homepage.");
			favouritesFactory.verifyFavoritesHeaderPresent(strTCID, strExpectedValue);

		}
	}

	@SuppressWarnings("static-access")
	@Test(priority = 15)
	public void verifyFavouriteMessageWhenNoFavouritePresent() throws Throwable {
		strTCID = "TC16_Homepage";
		strExpectedValue = BaseClass.getExcelDataByTestCaseID(strTCID);

		if (!strExpectedValue.equalsIgnoreCase("None")) {
			LogFactory.beginTestCase("Verify No Favourites message is displayed on Homepage.");
			favouritesFactory.verifyWhenNoFavouritePresent(strTCID, strExpectedValue);

		}
	}

	@SuppressWarnings("static-access")
	@Test(priority = 16)
	public void verifyFavouriteIconWhenNoFavouritePresent() throws Throwable {
		strTCID = "TC17_Homepage";
		strExpectedValue = BaseClass.getExcelDataByTestCaseID(strTCID);
		if (!strExpectedValue.equalsIgnoreCase("None")) {
			LogFactory.beginTestCase("Verify Favourites Star Icon with message on Homepage");
			favouritesFactory.verifyFavouriteIconWhenNoFavouriteAdded(strTCID, strExpectedValue);
		}

	}

	@SuppressWarnings("static-access")
	@Test(priority = 17)
	public void verify_homepagequicklink() throws Throwable {
		strTCID = "TC18_Homepage";
		strExpectedValue = BaseClass.getExcelDataByTestCaseID(strTCID);
		if (!strExpectedValue.equalsIgnoreCase("None")) {
			LogFactory.beginTestCase("Verify favourite Quick Links on HomePage");
			favouritesFactory.favouriteQuickLinkOnHomePage(strTCID);

		}

	}

	@SuppressWarnings("static-access")
	@Test(priority = 18)
	public void verifyQuickLinkFavouritesDropdown() throws Throwable {
		strTCID = "TC19_Homepage";
		strExpectedValue = BaseClass.getExcelDataByTestCaseID(strTCID);
		if (!strExpectedValue.equalsIgnoreCase("None")) {
			LogFactory.beginTestCase("Verify Favourites Quick Links Header");
			favouritesFactory.homePageQuickLinkContent(strTCID, strExpectedValue);

		}
	}

	/**
	 * @author shrey.choudhary
	 * @createdAt 22-05-2018
	 * @throws Throwable
	 *             modifiesAt 22-05-2018
	 */
	@Test(priority = 19)
	public void verifyDealerPrincipleRole() throws Throwable {
		
		String strTCID = "TC20_Homepage";
		String testData = getExcelDataByTestCaseID(strTCID);
	
		if (!strExpectedValue.equalsIgnoreCase("None")) {
			LogFactory.beginTestCase("Verify Dealer Principal role test case begins");
			homePageFactory.getDealerPrincipalRole(flagDealerPrinciple,testData,strTCID);
		}
	}

	@Test(priority = 20)
	public void verifyProductSegmentsOrder() throws Throwable {

		strTCID = "TC21_Homepage";

		strExpectedValue = getExcelDataByTestCaseID(strTCID);
		if (!strExpectedValue.equalsIgnoreCase("None")) {
			LogFactory.beginTestCase("Verify the available product segments are displayed in order");
			productSegmentFactory.checkOrderOfProductSegment(strExpectedValue, strTCID);
		}

	}

	@Test(priority = 21)
	public void verifyNotificationIconHomepage() throws Throwable {
		strTCID = "TC22_Homepage";
		LogFactory.beginTestCase(strTCID + "Verify Notification Icon test case begins");
		homePageFactory.verifyNotificationIconOnHomePage(strTCID);
		// LogFactory.endTestCase();
	}

	@Test(priority = 22)
	public void verifyNotificationList() throws Throwable {
		strTCID = "TC23_Homepage";
		strExpectedValue = getExcelDataByTestCaseID(strTCID);
		LogFactory.beginTestCase(strTCID + "Verify Notifications list test case begins");
		homePageFactory.verifyNotificationsList(strExpectedValue, strTCID);

	}

	@Test(priority = 23)
	public void verifyLeftNavigationWindow() throws Throwable {
		strTCID = "TC24_Homepage";
		strExpectedValue = getExcelDataByTestCaseID(strTCID);

		if (!strExpectedValue.equalsIgnoreCase("None")) {
		//	GenericFactory.navigateToHomePage();
			LogFactory.beginTestCase("Verify Portal Left Navigation Links Order");
			portalLeftNavFactory.compareNavigationLinksWithTestData(strExpectedValue, strTCID);
		}
	}

	@Test(priority = 24)
	public static void verifyAnnouncementFooterLinksOnHomepage() throws Throwable {
		strTCID = "TC25_Homepage";
		strExpectedValue = getExcelDataByTestCaseID(strTCID);
		LogFactory.beginTestCase(strTCID + "Verify Footer Links on Homepage.");
		homePageFactory.verifyFooterLinksonHomepage(strExpectedValue, strTCID);
	}
	
	@Test(priority = 25)
	public static void verifyCopyrightOfFooterLinksonHomepage() throws Throwable {
		strTCID = "TC26_Homepage";
		//strExpectedValue = getExcelDataByTestCaseID(strTCID);
		LogFactory.beginTestCase(strTCID + "Verify Copy right year on Footer Links on Homepage.");
		homePageFactory.verifyCopyrightOfFooterLinksonHomepage( strTCID);
	}
	
	@AfterClass
	public void getReportFooter() throws InterruptedException {
		LogFactory.endTestCase("Home Page Testcases");
		ReportFactory.tableEnd();

	}

}