package com.deere.TestCasesFactory;

import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.List;

import org.openqa.selenium.WebDriver;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import com.deere.Helpers.BaseClass;
import com.deere.Helpers.ExcelFactory;
import com.deere.Helpers.GenericFactory;
import com.deere.Helpers.LogFactory;
import com.deere.Helpers.ReportFactory;
import com.deere.PageFactory.Alerts_POF;

public class Alerts_TestCase extends BaseClass 
{
	WebDriver driver;
	static String strExpectedValue;  
	static String strTCID;  

	@BeforeClass
	public void getReportHeader() throws InterruptedException
	{
		ReportFactory.tableEnd();
		GenericFactory.createHeaderSection("Alerts");
	}

	/*
	 @Test(priority = 1) 
	 public static void verifyAlertsPortletContent() throws Throwable {
	  
		 List<LinkedHashMap> userWCMContent =
		 ExcelFactory.getWCMContentDetails("AT-Alert");
		 
		 System.out.println(">>>>>>>>>>>>>>>>>" + userWCMContent.size());
		 for (int i =0; i <userWCMContent.size(); i++) 
		 { 
			 String wcmTestCaseID = (String) userWCMContent.get(i).get("Test Case ID").toString().trim();
			 String userCountry=(String)userWCMContent.get(i).get("Dealer_Country"); 
			 String MRUCountry=(String)userWCMContent.get(i).get("Country");
			 String Title = (String) userWCMContent.get(i).get("Title");
			 String userProductType=(String) userWCMContent.get(i).get("Dealer_ProductType"); 
			 String wcmProductType=(String) userWCMContent.get(i).get("ProductType");
			 String Contenttype=(String) userWCMContent.get(i).get("ContentType");
			 		 
 		   alertPageFactory.verifyAlertsOnTheAlertPortlet(wcmTestCaseID,userCountry, MRUCountry, userProductType, wcmProductType,Contenttype, Title);

		 } 
	 }
	
	  @Test(priority = 2) 
		 public static void verifyAlertsContentFilteredOnProductType() throws Throwable {
		  
		 List<LinkedHashMap> userWCMContent =ExcelFactory.getWCMContentDetails("AT-Alert");
		
		 
		 	for (int i =0; i <userWCMContent.size(); i++) 
		 	{ 
		 		String wcmTestCaseID = (String) userWCMContent.get(i).get("Test Case ID").toString().trim();
		 		String userCountry=(String)userWCMContent.get(i).get("Dealer_Country"); 
		 		String wcmCountry=(String)userWCMContent.get(i).get("Country");
		 		String Title = (String) userWCMContent.get(i).get("Title");
		 		String userProductType=(String) userWCMContent.get(i).get("Dealer_ProductType"); 
		 		String wcmProductType=(String) userWCMContent.get(i).get("ProductType");
		 
		 		LogFactory.beginTestCase("verify alerts content filtered based on product types");
		 		alertPageFactory.checkAlertContentFilterOnProductTypes(userCountry,wcmCountry,userProductType, wcmProductType, Title, wcmTestCaseID);
		 	} 
		 }

	  @Test(priority = 2)
		public static void verifyAlertsPresenceOnHomePage() throws Throwable {

	 		LogFactory.beginTestCase("Alerts presence on Homepage ");
			strTCID = "TC01_ALERTS_HOMEPAGE";
			String strAlertHeaderTextFromExcel = BaseClass.getExcelDataByTestCaseID(strTCID);

			if (!strAlertHeaderTextFromExcel.equalsIgnoreCase("None")) {
				Alerts_POF.getAlertPortletHeaderAndText(strAlertHeaderTextFromExcel, "NA", strTCID);
			}
			strTCID = "TC02_ALERTS_HOMEPAGE";
			Alerts_POF.checkAlertIconDisplayedInHeader("NA", strTCID);
 
			strTCID = "TC03_ALERTS_HOMEPAGE";
			String strPortletHeaders= BaseClass.getExcelDataByTestCaseID(strTCID);
			Alerts_POF.checkOrderOfPortletsOnHomepage("NA", strTCID);
	 		 
			strTCID = "TC04_ALERTS_HOMEPAGE";
			if(!strTCID.equalsIgnoreCase("NA") || !strTCID.equals(" "))
			{
				System.out.println("date :"+BaseClass.getExcelDataByTestCaseID(strTCID));
				BaseClass.dateformat =  BaseClass.getExcelDataByTestCaseID(strTCID);
				
			Alerts_POF.checkAlertsInDescendingOrderOnDateTime(BaseClass.getExcelDataByTestCaseID("TC01_ALERTS_HOMEPAGE"), "NA", strTCID);
			}
			else {
				
				LogFactory.info("Specific date format is not mentioned in test data to verify");
			}   
		}
 */
	 	@Test(priority = 3)
		public static void verifyAlertsPresenceOnDepartments() throws Throwable {

			LogFactory.beginTestCase("Alerts presence on Departments ");
			strTCID = "TC05_ALERTS_DEPARTMENT";
			String strDepartmentNamesFromExcel = BaseClass.getExcelDataByTestCaseID(strTCID);
			
			if (!strDepartmentNamesFromExcel.equalsIgnoreCase("None"))
			{
				Alerts_POF.getAlertPortletHeaderAndText(BaseClass.getExcelDataByTestCaseID("TC01_ALERTS_HOMEPAGE"),strDepartmentNamesFromExcel, strTCID);
			}
			
			strTCID = "TC06_ALERTS_DEPARTMENT";
			Alerts_POF.checkAlertIconDisplayedInHeader(strDepartmentNamesFromExcel, strTCID);
			strTCID = "TC05_ALERTS_DEPARTMENT";
			Alerts_POF.checkOrderOfPortletsOnDepartment(strDepartmentNamesFromExcel,strTCID);
/*			strTCID = "TC08_ALERTS_DEPARTMENT";
			Alerts_POF.verifyAlertsShowingInDescendingOrderOnDateTime(BaseClass.getExcelDataByTestCaseID("TC01_ALERTS_HOMEPAGE"), strDepartmentNamesFromExcel, strTCID);
*/			
		}	  

	 	
		@Test(priority = 3)

		public static void verifyOrderOfAnnouncemnts() throws Throwable {

			LogFactory.beginTestCase("Verify whether the latest Alerts are showing at top on the Alerts portlet");

			strTCID = "TC08_ALERTS_DEPARTMENT";
			strExpectedValue = BaseClass.getExcelDataByTestCaseID(strTCID);
			if (!strExpectedValue.equalsIgnoreCase("None")) {

				alertPageFactory.verifyAlertsShowingInDescendingOrderOnDateTime(strTCID, strExpectedValue);

			}
		}

		 @Test(priority = 4)
	 	 
			public static void verifyAlertsHeaderFormat() throws Throwable {

				LogFactory.beginTestCase("Alerts format content on My Dealerpath ");
				strTCID = "TC09_ALERTS_HOMEPAGE";
				Alerts_POF.checkAlertHeaderContentWithFormat(strTCID);
				strTCID="TC10_ALERTS_HOMEPAGE";
				Alerts_POF.verifyEmbededlinks(strTCID);

			}
	 	
		 
	 
		 	@Test(priority = 5)
			public static void verifyReadMoreAndCollapseLinkOnAlerts() throws Throwable {
				LogFactory.beginTestCase("Verify the Readmore and Collapse links on the Alert Portlet");
				strTCID = "TC11_ALERTS_HOMEPAGE";
				Alerts_POF.checkReadMoreAndCollapseLinkForAlert(strTCID);
			}

	/*	 
	 	 @Test(priority = 1) 
	 		  public static void verifyAlertsOnIndexPage() throws Throwable {
	 			  
	 				 List<LinkedHashMap> userWCMContent =ExcelFactory.getWCMContentDetails("AT-Index Page");
	 				 
	 				 for (int i =0; i <userWCMContent.size(); i++) 
	 				 {
	 					
	 				 String wcmTestCaseID = (String) userWCMContent.get(i).get("Test Case ID").toString().trim();
	 				 String userCountry=(String)userWCMContent.get(i).get("Dealer_Country"); 
	 				 String wcmCountry=(String)userWCMContent.get(i).get("Country");
	 				 String Title = (String) userWCMContent.get(i).get("Title");
	 				 String userProductType=(String) userWCMContent.get(i).get("Dealer_ProductType"); 
	 				 List<String> userProductList = Arrays.asList(userProductType.split("(,)"));
	 				 
	 				 String wcmProductType=(String) userWCMContent.get(i).get("ProductType");
	 				 List<String> wcmProductList = Arrays.asList(wcmProductType.split("(,)"));
	 				 
	 				 String strDepartmentName=(String) userWCMContent.get(i).get("DepartmentName");
	 				 String strSecondlevelDepartment=(String) userWCMContent.get(i).get("2ndLevel");
	 				 String strIndexPage=(String) userWCMContent.get(i).get("3rdLevelIndexPage");
	 				 
	 				 String strThirdLevelChildIndexPage=(String) userWCMContent.get(i).get("3rdLevelChildIndexPage");
	 				 String strThirdLevelIndexFolder=(String) userWCMContent.get(i).get("3rdLevelFolder");
	 				 String strFourthLevelChildIndexPage=(String) userWCMContent.get(i).get("4thLevelChildIndexPage");
	 				 
	 				
	 				 LogFactory.beginTestCase("verify alerts presence");
	 				 ReportFactory.reportSectionName("Alerts_Index Page");
	 				 
	 				 if(strDepartmentName!="NA" && strSecondlevelDepartment!="NA" && strIndexPage!="NA")
	 				 {
	 					 
	 					 if(strThirdLevelChildIndexPage.equalsIgnoreCase("NA") && strThirdLevelIndexFolder.equalsIgnoreCase("NA") && strFourthLevelChildIndexPage.equalsIgnoreCase("NA"))
	 					 {
	 						LogFactory.info("Alert testcase for first index page");	
	 						 alertPageFactory.checkAlertsPresenceOnFirstIndexPage(userCountry, wcmCountry, userProductList, wcmProductList, strDepartmentName, strSecondlevelDepartment, strIndexPage,wcmTestCaseID);
	 					 }
	 					 else {
	 						 
	 						 LogFactory.info("Data is not provided as expected for first level index page");
	 						 String strFlag="Pass";
	 						 String strResult="Data is not provided as expected for first level index page";
	 						 ReportFactory.reporterOutput(strTCID,"Verify alerts portlet not available on index pages", "Department name :" +strDepartmentName +" and secondlevel department is: "+strSecondlevelDepartment,"Alerts portlet should not be present on index pages",strResult, strFlag);  
	 					 }
	 					 		 
	 				 }
	 				 else {
	 					 
	 					 LogFactory.info("Department, or second level department or index page data is not mentioned in the sheet" );
	 					 String strFlag="Pass";
	 					 String strResult="Department, or second level department or index page data is not mentioned in the sheet";
	 					 ReportFactory.reporterOutput(strTCID,"Verify alerts portlet not available on index pages", "Department name :" +strDepartmentName +" and secondlevel department is: "+strSecondlevelDepartment,"Alerts portlet should not be present on index pages",strResult, strFlag);  
	 					 
	 				 }
	 				 
	 				 }
	 				 
	 		  }
	 		
	*/ 		
	  
	  
	
	  @AfterClass
		public void getReportFooter() throws InterruptedException {
			LogFactory.endTestCase("Alerts Testcases");
			ReportFactory.tableEnd();

		}

	
	
	
	
	
	
	
	
	
	
	
	
	
	

}
