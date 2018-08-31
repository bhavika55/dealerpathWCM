package com.deere.TestCasesFactory;


import java.util.LinkedHashMap;
import java.util.List;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import com.deere.Helpers.BaseClass;
import com.deere.Helpers.ExcelFactory;
import com.deere.Helpers.GenericFactory;
import com.deere.Helpers.LogFactory;
import com.deere.Helpers.ReportFactory;

@Test(groups = { "PortletLinks" })
public class PortletLinks_TestCase extends BaseClass {

	WebDriver driver;
	static String strExpectedValue;
	static String strTCID;

	@BeforeClass
	public void getReportHeader() throws InterruptedException {
		ReportFactory.tableEnd();
		GenericFactory.createHeaderSection("Portlet Links");
	}

	@Test(priority = 1)
	public void VerifyPortletlinks() throws Throwable {
		String UserDefinedCountry = "";
		String WCMCountry = "";
		String UserDefinedProducts = "";
		String WCMProducts = "";
		String RootsiteArea = "";
		String Contenttype = "";
		String DepartmentName = "";
		String level2 = "";
		String Title = "";
		String wcmTestCaseID="";

		List<LinkedHashMap> userWCMContent = ExcelFactory.getWCMSiteAreaDetails("NA");

		for (int i = 0; i <userWCMContent.size(); i++) {
			
			wcmTestCaseID = (String) userWCMContent.get(i).get("Test Case ID").toString().trim();
			UserDefinedCountry = (String) userWCMContent.get(i).get("Dealer_Country").toString().trim();
			WCMCountry = (String) userWCMContent.get(i).get("Country").toString().trim();
			UserDefinedProducts = (String) userWCMContent.get(i).get("Dealer_ProductType").toString().trim();
			WCMProducts = (String) userWCMContent.get(i).get("ProductType").toString().trim();
			//RootsiteArea = (String) userWCMContent.get(i).get("RootSiteArea").toString().trim();
			Contenttype = (String) userWCMContent.get(i).get("ContentType").toString().trim();
			DepartmentName = (String) userWCMContent.get(i).get("DepartmentName").toString().trim();
			level2 = (String) userWCMContent.get(i).get("2ndLevel").toString().trim();
			Title = (String) userWCMContent.get(i).get("Title").toString().trim();
			// System.out.println(title+"Test excel Data@@@@@@@@"+rst);
	
			portletLinksFactory.PortletLinks(wcmTestCaseID,UserDefinedCountry, WCMCountry, UserDefinedProducts, WCMProducts, RootsiteArea,
			Contenttype, DepartmentName, level2, Title);
				
			//	LogFactory.endTestCase();

		}
	}

	@AfterClass
	public void getReportFooter() throws InterruptedException {
		LogFactory.endTestCase("Home Page Testcases");
		//ReportFactory.tableEnd();
	}
}