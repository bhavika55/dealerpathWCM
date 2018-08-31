package com.deere.TestCasesFactory;

import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.List;

import org.openqa.selenium.WebDriver;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import com.deere.Helpers.BaseClass;
import com.deere.Helpers.ExcelFactory;
import com.deere.Helpers.GenericFactory;
import com.deere.Helpers.LogFactory;
import com.deere.Helpers.ReportFactory;
import com.deere.PageFactory.Alerts_POF;
import com.deere.PageFactory.Announcements_POF;

public class Announcement_TestCase extends BaseClass {

	WebDriver driver;
	static String strExpectedValue;
	static String strTCID;

	@BeforeClass
	public void getReportHeader() throws InterruptedException {
		ReportFactory.tableEnd();
		GenericFactory.createHeaderSection("Announcement");
	}

	@Test(priority = 1)
	public static void verifyAlertsContentFilteredOnProductType() throws Throwable {

		@SuppressWarnings("rawtypes")
		List<LinkedHashMap> userWCMContent = ExcelFactory.getWCMContentDetails("AT-Announcement");

		for (int i = 0; i < userWCMContent.size(); i++) {

			String wcmTestCaseID = (String) userWCMContent.get(i).get("Test Case ID").toString().trim();
			String userCountry = (String) userWCMContent.get(i).get("Dealer_Country").toString().trim();
			String MRUCountry = (String) userWCMContent.get(i).get("Country").toString().trim();
			String Title = (String) userWCMContent.get(i).get("Title").toString().trim();
			String userProductType = (String) userWCMContent.get(i).get("Dealer_ProductType").toString().trim();
			String wcmProductType = (String) userWCMContent.get(i).get("ProductType").toString().trim();
			String contenttype = (String) userWCMContent.get(i).get("ContentType");
			String copyToDepartment = (String) userWCMContent.get(i).get("CopyToDepartment");
			String wcmDealerType = (String) userWCMContent.get(i).get("DealerType (Main/Sub)").toString().trim();
			String departmentName = (String) userWCMContent.get(i).get("DepartmentName").toString().trim();

			LogFactory.beginTestCase("verify Announcement content on Announcement Portlet");
			announcementFactory.verifyAnnouncementsOnTheAnnouncementPortlet(wcmTestCaseID, userCountry, MRUCountry,
					departmentName, copyToDepartment, wcmDealerType, userProductType, wcmProductType, contenttype,
					Title);

		}
	}

	@Test(priority = 2)
	public static void verifyAnnouncementOnDepartmentUncheck() throws Throwable {
		String wcmDeptName = "";
		String titleName = "";
		String strTCID = "";
		String wcmCopyToDepartment = "";
		List<LinkedHashMap> userWCMContent = ExcelFactory
				.getUserWcmDetailsAfterFilteringCountryAndProduct("AT-Announcement");
		
		GenericFactory.navigateToHomePage();
		for (int i = 0; i <= userWCMContent.size(); i++) {

			wcmDeptName = (String) userWCMContent.get(i).get("DepartmentName").toString().trim();
			titleName = (String) userWCMContent.get(i).get("Title").toString().trim();
			strTCID = (String) userWCMContent.get(i).get("Test Case ID").toString().trim();
			wcmCopyToDepartment = (String) userWCMContent.get(i).get("CopyToDepartment").toString().trim();
			
			announcementFactory.verifyAnnouncementTitleOnDepartmentUncheck(wcmDeptName, titleName, strTCID,
					wcmCopyToDepartment);
		}

	}
	
	@Test(priority = 3)
	public static void verifyAnnouncementOnProductUncheck() throws Throwable {
		String UserDefinedProducts = "";
		String WCMProducts = "";
		String Title = "";
		String wcmTestCaseID = "";
		List<LinkedHashMap> userWCMContent = ExcelFactory
				.getUserWcmDetailsAfterFilteringCountryAndProduct("AT-Announcement");

		System.out.println("userWCMContent===========" + userWCMContent);
		
		for (int i = 0; i < userWCMContent.size(); i++) {

			wcmTestCaseID = (String) userWCMContent.get(i).get("Test Case ID").toString().trim();
			UserDefinedProducts = (String) userWCMContent.get(i).get("Dealer_ProductType").toString().trim();
			WCMProducts = (String) userWCMContent.get(i).get("ProductType").toString().trim();
			Title = (String) userWCMContent.get(i).get("Title").toString().trim();

			announcementFactory.verifyAnnouncementTitleOnProductUncheck(UserDefinedProducts, WCMProducts, Title,
					wcmTestCaseID);
		}
	}


	
	@Test(priority = 4)

	public static void verifyAnnouncementHeaderFormat() throws Throwable {

		LogFactory.beginTestCase("Verifying the Announcement header format");

		strTCID = "TC09_Announcement_Portlet";
		strExpectedValue = BaseClass.getExcelDataByTestCaseID(strTCID);
		if (!strExpectedValue.equalsIgnoreCase("None")) {

			announcementFactory.verifyAnnouncementHeaderContentWithFormat(strTCID);

		}

	}

	@Test(priority = 5)

	public static void verifyAnnouncementPortletEmbededLinks() throws Throwable {

		LogFactory.beginTestCase("Verifying the embeded links showing on the Announcement portlet");

		strTCID = "TC10_Announcement_Portlet";
		strExpectedValue = BaseClass.getExcelDataByTestCaseID(strTCID);
		if (!strExpectedValue.equalsIgnoreCase("None")) {

			announcementFactory.verifyAnnouncementEmbededlinks(strTCID);

		}
	}


	@Test(priority = 6)

	public static void verifyOrderOfAnnouncemnts() throws Throwable {

		LogFactory.beginTestCase("Verify whether the latest Announcements are showing at top on the Announcement portlet");

		strTCID = "TC11_Announcement_Portlet";
		strExpectedValue = BaseClass.getExcelDataByTestCaseID(strTCID);
		if (!strExpectedValue.equalsIgnoreCase("None")) {

			announcementFactory.verifyAnnouncemntsShowingInDescendingOrderOnDateTime(strTCID, strExpectedValue);

		}
	}

	@Test(priority = 7)

	public static void verifyAnnouncementCountOnHeader() throws Throwable {
		LogFactory.beginTestCase("Verifying the announcement count showing on Announcement portlet header");
		strTCID = "TC12_Announcement_Portlet";
		strExpectedValue = BaseClass.getExcelDataByTestCaseID(strTCID);
		
		if (!strExpectedValue.equalsIgnoreCase("None")) {
			announcementFactory.verifyAnnouncementHeaderCount(strTCID);
		}
	}	
	

	@Test(priority = 8)

	public static void verifyReadMoreAndCollapseLinkOnAnnouncement() throws Throwable {
		LogFactory.beginTestCase("Verify the Readmore and Collapse links on the Announcement Portlet");
		strTCID = "TC13_Announcement_Portlet";
		strExpectedValue = BaseClass.getExcelDataByTestCaseID(strTCID);
		
		if (!strExpectedValue.equalsIgnoreCase("None")) {
			announcementFactory.checkReadMoreAndCollapseLinkForAnnouncement(strTCID);
		}
	}	
	
	@Test(priority = 9)

	public static void verifyExpandAndCollapseLinkOnAnnouncement() throws Throwable {
		LogFactory.beginTestCase("Verify the Expand and Collapse links on the Announcement Portlet");
		strTCID = "TC14_Announcement_Portlet";
		strExpectedValue = BaseClass.getExcelDataByTestCaseID(strTCID);
		
		if (!strExpectedValue.equalsIgnoreCase("None")) {
			announcementFactory.verifyExpandAndCollapseLinkOnAnnuoncementPortlet(strTCID);
		}
	}	
	

	
}
