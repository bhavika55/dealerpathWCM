//Create separate class under page factory with name announcement_Factory and pase it u can run this method any where by classname.methodname

package com.deere.PageFactory;

import static org.testng.Assert.fail;

import java.net.HttpURLConnection;
import java.net.URL;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.TreeMap;
import java.util.regex.Pattern;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.How;
import org.openqa.selenium.support.ui.Select;
import org.testng.Assert;

import com.deere.Helpers.BaseClass;
import com.deere.Helpers.ExcelFactory;
import com.deere.Helpers.GenericFactory;
import com.deere.Helpers.LogFactory;

import com.deere.Helpers.ReportFactory;
import com.deere.Helpers.ValidationFactory;

public class Announcements_POF extends BaseClass {

	final WebDriver AnncmntDriver;

	public Announcements_POF(WebDriver driver) {

		this.AnncmntDriver = driver;

	}

	@FindBy(how = How.XPATH, using = ".//div[@class='section']//div[@class='list-item-body']/span")
	public static List<WebElement> AnnouncementTitleContents;

	@FindBy(how = How.XPATH, using = ".//*[@class='section-title announcement']")
	static WebElement wbelHeaderTitleAnnouncement;

	@FindBy(how = How.XPATH, using = ".//*[@class='wide-list hide-overflow is-expanded']")
	static WebElement wbelBodyAnnouncement;

	@FindBy(how = How.XPATH, using = ".//*[@class='wide-list hide-overflow is-expanded']/div")
	static List<WebElement> listAnnouncementBody;

	@FindBy(how = How.XPATH, using = ".//div/div/div/div/section/div/div/div/h3")
	static List<WebElement> listHeader;

	@FindBy(how = How.XPATH, using = ".//*[@class='wide-list hide-overflow is-expanded']/div/div[1]")
	static List<WebElement> listAnnouncementTitleBody;

	@FindBy(how = How.XPATH, using = ".//*[@id='layoutContainers']/div/div[2]/div[3]/div[3]/section/div/div[2]/div")
	static List<WebElement> listAnnouncementTableElements;

	@FindBy(how = How.XPATH, using = ".//div[@class='section']//div[@class='list-item-body']/span")
	static List<WebElement> listAnnouncementTableTitle;

	@FindBy(how = How.XPATH, using = ".//*[@class='list-item-info parse-department']")
	static List<WebElement> listAnnouncementDeptName;

	@FindBy(how = How.XPATH, using = ".//*[@class='leftNav']/li/a")
	static List<WebElement> listAnnouncementLeftNavigationDeptName;

	@FindBy(how = How.XPATH, using = ".//*[@id='layoutContainers']/div/div[2]/div[3]/div[3]/section/div/div[2]")
	public static WebElement wbelAnnouncementFramePath;
	
	@FindBy(how = How.XPATH, using = ".//div/div/div/div/section/div/div/div/h3")
    public static List<WebElement> HeaderList;

	@FindBy(how = How.XPATH, using = "//*[@class='group-value checkbox-value']//label[@class='click-target-only']//input[@type='checkbox']")
    public static List<WebElement> allCheck;

    @FindBy(how = How.XPATH, using = "//div[@class='group-value checkbox-value']/div[@class='value']")
    public static List<WebElement> allCheck1;

    @FindBy(how = How.XPATH, using = ".//div[@class='section']//div[@class='list-item-body']/span")
    public static List<WebElement> AnnouncementTableTitle;
    
    @FindBy(how = How.XPATH, using = ".//*[@id='js-segments-popover']/div[2]/div[3]/button")
    public static WebElement ApplyFilterButton;
    
	@FindBy(how = How.XPATH, using = "//div[@class='wide-list hide-overflow is-expanded']")
	public static WebElement Announcementportlet;
    
	@FindBy(how = How.XPATH, using = ".//*[@id='layoutContainers']/div/div[2]/div[3]/div[3]/section/div/div[2]/div[2]/div[2]/div")
	 public static WebElement linkAnnouncementOfExpand;
	
	
	public static boolean verifyAnnouncementTableOnHomePage() throws Throwable {
		try {
			List<WebElement> listAnnTableObj = listAnnouncementTableElements;

			if (listAnnTableObj.size() > 0) {

				return true;
			} else {
				return false;
			}
		} catch (Exception e) {
			// LogFactory.error("e");
			return false;
		}

	}

	// ****************************************************** E-O-M
	// **********************************************************************************

	public static boolean verifyAnnouncementTableBodyIsprsentOnHomePage() throws Throwable {
		try {
			if (listAnnouncementTableElements.size() > 0) {

				if (ValidationFactory.isElementPresent(wbelBodyAnnouncement)) {
					return true;

				} else {

					return false;
				}
			} else {

				return false;
			}

		}

		catch (Exception e) {
			// LogFactory.error("e");
			return false;
		}
	}

	// ********************************************************** E- O - M
	// ****************************************************************

	// To verify announcement header is present or not
	public static boolean verifyAnnouncementTableHeaderIsprsentOnHomePage() throws Throwable {
		try {
			if (listAnnouncementTableElements.size() > 0) {

				if ((ValidationFactory.isElementPresent(wbelBodyAnnouncement))) {

					return true;

				} else {
					return false;

				}

			} else {

				return false;
			}

		}

		catch (Exception e) {
			// LogFactory.error("e");
			return false;
		}
	}

	// ****************************************************** E-O-M
	// **********************************************************************************

	// To verify announcement header text in preferred language

	public static String verifyAnnouncementHeaderTextPrefferedLang() throws Throwable {
		String strActualHeaderText = null;
		try {
			if (listAnnouncementTableElements.size() > 0) {
				// LogFactory.info("To verify announcement header text in preferred language ");

				if (ValidationFactory.isElementPresent(wbelHeaderTitleAnnouncement)) {
					String strHeaderTitle = wbelHeaderTitleAnnouncement.getText().toString();
					strActualHeaderText = strHeaderTitle.substring(0, strHeaderTitle.lastIndexOf('(') - 1);
					return strActualHeaderText;

				} else {
					return "";

				}
			} else {
				// LogFactory.info("Announcement portlet is not Present.");
				return "";
			}
		}

		catch (Exception e) {
			LogFactory.error("e");
			return "";
		}
	}

	/**
	 * @param TCID
	 * @throws Throwable
	 */
	// To verify announcement Table is present or not
	public static void verifyAnnouncementTableOnHomePage(String strTCID) throws Throwable {
		String strFlag = "Fail";
		String strResult = "Announcement portlet is not displaying.";

		try {
			List<WebElement> listAnnTableObj = listAnnouncementTableElements;

			if (listAnnTableObj.size() > 0) {
				strFlag = "Pass";
				strResult = "Announcement portlet is displaying.";

				LogFactory.info("Announcement portlet is present");

			}

			ReportFactory.reporterOutput(strTCID, "Verify Announcement portlet is present on Home Page", "NA",
					"Announcement should displayed.", strResult, strFlag);

			if (strFlag.equalsIgnoreCase("FAIL")) {
				Assert.assertFalse(true);
			}

		} catch (Exception e) {
			// LogFactory.error("e");
			String er = e.getMessage().toString();

			ReportFactory.reporterOutput(strTCID, "Verify Announcement portlet is present on Home Page", "NA",
					"Announcement portlet should be displayed.", strResult, er.substring(0, 25));

		}
	}

	// ********************************************************** E- O - M
	// ****************************************************************

	// To verify announcement body is present or not
	/**
	 * @param TCID
	 * @throws Throwable
	 */
	public static void verifyAnnouncementContentIsPrsent(String strTCID) throws Throwable {
		String strFlag = "Fail";
		String strResult = "Announcement portlet content/body is not present";

		try {
			List<WebElement> listAnnTableObj = listAnnouncementTableElements;

			if (listAnnTableObj.size() > 0) {

				strFlag = "Pass";
				strResult = "Announcement portlet content/body is present";

				/*
				 * // LogFactory.info("Announcement is present"); if
				 * (Announcements_POF.verifyAnnouncementTableBodyIsprsentOnHomePage()) { strFlag
				 * = "Pass"; strResult="Announcement portlet content/body is present";
				 * LogFactory.info("Announcement portlet is present"); }
				 * 
				 * if (Announcements_POF.verifyAnnouncementTableHeaderIsprsentOnHomePage()) {
				 * strFlag = "Pass"; strResult="Announcement content/body is present";
				 * 
				 * LogFactory.info("Announcement header is present");
				 * 
				 * }
				 */
			}

			ReportFactory.reporterOutput(strTCID, "Verify Announcement content/body is showing", "NA",
					"Announcement content/body should be display", strResult, strFlag);

		} catch (Exception e) {
			String er = e.getMessage().toString();

			ReportFactory.reporterOutput(strTCID, "Verify Announcement content/body is showing", "NA",
					"Announcement Content/body should be dispaly", strResult, er.substring(0, 25));
		}
	}

	// **************************************************** E-O-M
	// ****************************************************************************

	// To verify announcement header text in preferred language

	/**
	 * @param expectedHeaderTxt
	 * @param TCID
	 * @throws Throwable
	 */
	public static void verifyAnnouncementHeaderTextPrefferedLang(String strExpectedHeaderTxt, String strTC_ID)
			throws Throwable {

		String strTCID = "TC12_Homepage";
		String strActualHeaderText = null;
		String strFlag = "Fail";
		String strResult = "Announcement header is showing as per the user preferred language.";

		try {
			if (listAnnouncementTableElements.size() > 0) {

				// LogFactory.info("Verify announcement header is showing as per the user
				// preferred language. ");

				if (ValidationFactory.isElementPresent(wbelHeaderTitleAnnouncement)) {

					String txtheadertxt = wbelHeaderTitleAnnouncement.getText().toString().trim();
					String[] headertxt = txtheadertxt.split(Pattern.quote("("));

					String actualHeaderTxt = headertxt[0].trim();
					if (actualHeaderTxt.length() > 0) {

						if (actualHeaderTxt.equals(strExpectedHeaderTxt)) {
							strFlag = "Pass";

							LogFactory.info("Announcement header is showing as per the user preferred language."
									+ "Actual Value-" + strActualHeaderText + ";" + "Expected Value-"
									+ strExpectedHeaderTxt);
						}
					}
				}
			}
			ReportFactory.reporterOutput(strTCID,
					"Verify Announcement header text is Present as per the user preferred language.",
					strExpectedHeaderTxt, "Announcement header should show as per the user preferred language.",
					strResult, strFlag);

			/*
			 * if (strFlag.equalsIgnoreCase("FAIL") ) { Assert.assertFalse(true);}
			 */

		} catch (Exception e) {
			// LogFactory.error("e");
			String er = e.getMessage().toString();

			ReportFactory.reporterOutput(strTCID,
					"Verify Announcement header text is showing as per the user preferred language.", "NA",
					"Announcement Content should dispaly", strResult, er.substring(0, 25));
		}
	}

	// ****************************************************** E-O-M
	// ******************************************************************************

	// verify Announcement in detail
	public static void verifyAnnouncementInDetail_Homepage() throws Throwable {

		String strTCID = "TC14_Announcement";
		String strFlag = "Fail";
		String strresult = "Announcement is displaying in Detail on Homepage";

		try {
			// To get all title of Announcement
			List<WebElement> listItemTitleDealerHomePage = listAnnouncementTableElements;
			List<WebElement> listItemInfoDepartment = listAnnouncementDeptName;

			for (int i = 0; i < listItemTitleDealerHomePage.size(); i++) {
				String titleValueDealerHomePage = listItemTitleDealerHomePage.get(i).getText().toString().trim();

				// To get Dept name
				String announcementDeptName = listItemInfoDepartment.get(i).getText().toString().trim();

				LogFactory.info("Navigating to department Page- " + announcementDeptName);

				// To get all dept list
				List<WebElement> deptListNames = listAnnouncementLeftNavigationDeptName;
				for (int k = 0; k < deptListNames.size(); k++) {
					String deptName = deptListNames.get(k).getText().toString().trim();
					if (deptName.equals(announcementDeptName)) {
						System.out.println(deptName + " = equals to = " + announcementDeptName);
						deptListNames.get(k).click();

						// Retrive and Compare
						List<WebElement> listItemTitleDeptPage = listAnnouncementTableTitle;

						for (int j = 0; j < listItemTitleDeptPage.size(); j++) {
							String titleValueDeptPage = listItemTitleDeptPage.get(j).getText().toString().trim();

							// to validate on desired dept
							if (titleValueDealerHomePage.equals(titleValueDeptPage)) {
								strFlag = "Pass";

								LogFactory.info(titleValueDealerHomePage
										+ " Homepage Announcement Deptartment Page Title is same as showing on DelerPath Homepage. "
										+ titleValueDeptPage);

								ReportFactory.reporterOutput(strTCID, "Verify Announcement In Detail Homepage",
										titleValueDealerHomePage, "Announcement is displaying in Detail on Homepage",
										strresult, strFlag);

							}

							else {

								continue;
							}
						}

						// to go back again to Announcement page
						// BaseClass.wbDv.navigate().back();

						BaseClass.wbDriver.findElement(By.xpath(".//*[@class='app-title']")).click();

						LogFactory.info("Navigating back to the DealerPath Homepage");

						ReportFactory.reporterOutput(strTCID, "Verify the Announcement In Detail on Homepage", "NA",
								"Announcement should show in Detail on Homepage", strresult, strFlag);

					}
				}

			}
		} catch (Exception e) {
			LogFactory.info(e.getMessage());
		}
	}

	// ****************************************************** E-O-M
	// ******************************************************************************

	public static void verifyDateFormat(String strTCID, String strHeader, String strDateFormat) throws Throwable {
		// String TCID = "TC14_Announcement";
		String strFlag = "Fail";
		String strResult = "Announcement date is not showing as per the UI Format.";

		try {
			List<String> listDateList = GenericFactory.getListOfDatesByHeaderName(strHeader);
			LogFactory.info("dateList= " + listDateList);

			// validate format
			boolean dateFormatValue = GenericFactory.isValidDateFormat(strDateFormat, listDateList);
			if (dateFormatValue) {
				LogFactory.info("Date format is right.");
			} else {
				LogFactory.info("Date format is right.");
			}
		} catch (Exception e) {
			// LogFactory.error("e");
			// String er = e.getMessage().toString().trim();

			ReportFactory.reporterOutput(strTCID, "verify Date Format", "NA",
					"Announcement Date format should display in yyyy-mm-dd. ", strResult, strFlag);
		}

	}

	// ***************************************************** E-O-M
	// *****************************************************************************

	// To do the date sort
	public static void verifyDateOrder(String strHeader) throws Throwable {
		String strTCID = "TC14_Announcement";
		String strFlag = "Fail";
		String strResult = "Announcement date is not showing as per the UI Format.";

		try {
			List<String> dateList = GenericFactory.getListOfDatesByHeaderName(strHeader);
			LogFactory.info("dateList= " + dateList);

			// sort
			boolean sortedDate = GenericFactory.verifyDateSortedOrder(dateList);
			if (sortedDate) {
				strFlag = "Pass";
				LogFactory.info("Date is in sorted order.");

				ReportFactory.reporterOutput(strTCID, "verify Date Order", "NA",
						"Announcement Date order should display in new to old. ", strResult, strFlag);
			} else {
				LogFactory.info("Date is not in sorted order.");

			}
		} catch (Exception e) {
			// LogFactory.error("e");
			// String er = e.getMessage().toString().trim();

			ReportFactory.reporterOutput(strTCID, "verify Date Order", "NA",
					"Announcement Date order should display in new to old. ", strResult, strFlag);
		}
	}
	// ***************************************************** E-O-M
	// *****************************************************************************

	// &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

	public static void verifyAnnouncementsOnTheAnnouncementPortlet(String strWCMTCID, String userDefinedCountry,
			String wcmCountry, String departmentName, String copyToDepartment, String wcmDealerType,
			String userDefinedProducts, String wcmProducts, String contentType, String announcementTitle)
			throws Throwable {

		boolean country = GenericFactory.userAndWCMCountryComparison(userDefinedCountry, wcmCountry);
		boolean product = GenericFactory.userAndWCMProductTypeComparison(userDefinedProducts, wcmProducts);
		boolean flagTitle = true;
		boolean booleanDealerType = false;
		List<String> departmentList = new ArrayList<String>();
		List<String> ListAllActiveLinks = new ArrayList<String>();

		String inputValue = "<B>WCM Maping :</B> Country : " + wcmCountry + " ,<B> Product Type : </B>  " + wcmProducts 
				+ "<B> Department  :</B>  " + departmentName + "<B> Copy to Department  :</B>  " + copyToDepartment 				
				+ " <B>Dealer Maping :</B> Country : " + userDefinedCountry + " ,Product Type : " + userDefinedProducts 
				+ "<B>Department Name :</B> " + departmentName;
		String testCaseDescription = " Verify the Announcement titles on the Announcement portlet for the Content Type "
				+ contentType;
		String expectedValue = "Verify the announcement " + announcementTitle
				+ " published and displayed in Announcement Portlet ";
		String actualValue = "";

		// Getting list of all active links from left navigation window
		if (!departmentName.equalsIgnoreCase("General")) {
			WebElement department = GenericFactory.getDeptname(departmentName);

			if (department != null) {
				LogFactory.info(" Clicking on the active department " + departmentName + " from homepage");
				department.click();
			} else {
					GenericFactory.clickOnFlyOutDeptMenu(departmentName);
			}
			flagTitle = verifyAnnouncementHeader(announcementTitle);
			
		} else {
			GenericFactory.navigateToHomePage();
			flagTitle = verifyAnnouncementHeader(announcementTitle);
		}

		if (country && product && flagTitle) {
			// compare DealerType with WCM DealerType content of AT-Announcement row.
			if (!flagDealerType.equalsIgnoreCase("NA")) {

				booleanDealerType = GenericFactory.verifyDealerType(flagDealerType, wcmDealerType);

				if (booleanDealerType) {

					if (!copyToDepartment.equalsIgnoreCase("NA")) {

						departmentList.addAll(GenericFactory.splitString(copyToDepartment, ","));
						for (int j = 0; j < departmentList.size(); j++) {

							WebElement department = GenericFactory.getDeptname(departmentList.get(j));

							if (department != null) {
								LogFactory.info(" Clicking on the active copy to department " + departmentList.get(j)
										+ " from homepage");
								department.click();
							} else {
								GenericFactory.clickOnFlyOutDeptMenu(departmentList.get(j));

							}
							if (verifyAnnouncementHeader(announcementTitle)) {

								inputValue = inputValue + "<B>Copy to Department Name :</B> "
										+ copyToDepartment;
								actualValue = actualValue
										.concat("<B>Title name :</B>  "
												+ announcementTitle + " is present in <B>" + departmentName + " and "
												+ departmentList.get(j) + " departments</B>");
								// GenericFactory.navigateToHomePage();
							} else {
									inputValue = inputValue + "<B>Copy to Department Name :</B> "
										+ copyToDepartment;
									actualValue = actualValue
										.concat(" <B>Title name :</B>  "
												+ announcementTitle + " is present in <B>" + departmentName
												+ "department but NOT present in " + departmentList.get(j)
												+ " department</B>");
								flagTitle = false;
								// GenericFactory.navigateToHomePage();
								break;
							}
						}
					} else {
							inputValue = inputValue + "<B> User Dealer Type :</B> " + BaseClass.flagDealerType
								+ "<br/>" + "<B> WCM DealerType :</B> " + wcmDealerType;
							actualValue = actualValue.concat(
								" <B>Title name :</B>  " + announcementTitle
										+ " is present in <B>" + departmentName + " department</B>");
						}
					} else {
					// which means WCM content dealer type = NA
				}
			}
			else {
				if (!copyToDepartment.equalsIgnoreCase("NA")) {
				departmentList.addAll(GenericFactory.splitString(copyToDepartment, ","));
				for (int j = 0; j < departmentList.size(); j++) {
					WebElement department = GenericFactory.getDeptname(departmentList.get(j));
					if (department != null) {
						LogFactory.info(" Clicking on the active copy to department " + departmentList.get(j)
								+ " from homepage");
						department.click();
					} else {
						GenericFactory.clickOnFlyOutDeptMenu(departmentList.get(j));
					}
					if (verifyAnnouncementHeader(announcementTitle)) {
						inputValue = inputValue + "<B>Copy to Department Name :</B> " + copyToDepartment;
						actualValue = actualValue
								.concat("<ul style=\"list-style-type:circle\"> <li> <B>Title name :</B>  "
										+ announcementTitle + " is present in <B>" + departmentName + " and "
										+ departmentList.get(j) + " departments</B></li></ul>");
					} else {
						inputValue = inputValue + "<B>Copy to Department Name :</B> " + copyToDepartment;
						actualValue = actualValue.concat(
								" <ul style=\"list-style-type:circle\"> <li> <B>Title name :</B>  " + announcementTitle
										+ " is present in <B>" + departmentName + "department but NOT present in "
										+ departmentList.get(j) + " department</B></li></ul>");
						flagTitle = false;
						break;
					}
				}
			} else {
				actualValue = actualValue.concat("<ul style=\"list-style-type:circle\"> <li> <B>Title name :</B>  "
						+ announcementTitle + " is present in <B>" + departmentName + " department</B></li></ul>");
			}
		}
		String flag = flagTitle ? "Pass" : "Fail";
		ReportFactory.reporterOutput(strWCMTCID,testCaseDescription, inputValue, expectedValue,actualValue, flag);
		}
		else if (country && product && !flagTitle) {
			ReportFactory.reporterOutput(strWCMTCID, testCaseDescription, inputValue, expectedValue,
					actualValue + "User product and country match but title is not present","Fail");
		} else if (country && !product && !flagTitle) {
			ReportFactory.reporterOutput(strWCMTCID, testCaseDescription, inputValue, expectedValue,
					actualValue+ "User product doesn't match hence title is not present","Pass");
		} else if (!country && !product && !flagTitle) {
			ReportFactory.reporterOutput(strWCMTCID,testCaseDescription, inputValue, expectedValue,
					actualValue+ "User product & country doesn't match hence title is not present","Pass");
		} else if (!country && product && flagTitle) {
			ReportFactory.reporterOutput(strWCMTCID,testCaseDescription, inputValue, expectedValue,
					actualValue + "User country doesn't match but title is present","Fail");
		} else if (!country && product && !flagTitle) {
			ReportFactory.reporterOutput(strWCMTCID,testCaseDescription, inputValue, expectedValue,
					actualValue	+ "User country doesn't match hence title is not present","Pass");
		} else if (!country && !product && flagTitle) {
			ReportFactory.reporterOutput(strWCMTCID,testCaseDescription, inputValue, expectedValue,
					actualValue	+ "User country & product doesn't match but title is present","Fail");
		} else if (country && !product && flagTitle) {
			ReportFactory.reporterOutput("Announcement title",testCaseDescription, inputValue,
					expectedValue,actualValue+ "User product doesn't match but title is present","Fail");
		}

	}

/*	public static boolean verifyAnnouncementTitleOnDepartment(String UserDeptName, String titleName) throws Throwable {

		if (!UserDeptName.equalsIgnoreCase("General")) {
			
			if (GenericFactory.clickOnDepartmentByName(UserDeptName)) {
				if (verifyAnnouncementHeader(titleName)) 
					return true;
				}
		} else {
			if (verifyAnnouncementHeader(titleName)) 	
				return true;
		}
		LogFactory.info("Announcement title "+ titleName +" is not present on the announcement portlet");
		return false;
	}*/

	
	
	public static boolean verifyAnnouncementHeader(String Title) throws Throwable {
		try {
			List<WebElement> listItemTitleDeptPage = AnnouncementTitleContents;
		//	List<String> listAnnouncementTitle = new ArrayList<String>();
			for (int j = 0; j < listItemTitleDeptPage.size(); j++) {
				String titleValueOfDeptPage = listItemTitleDeptPage.get(j).getText().toString().trim();
				
				if (titleValueOfDeptPage.contains(Title)) {
					System.out.println("test");
					return true;
			
				
				}
				
			}
			return false;
		} catch (Exception e) {
			e.getMessage();
			return false;
		}
	
	}
	
	
	public static void verifyAnnouncementHeaderContentWithFormat(String strTCID) throws Throwable {
		
		String strFlag = "Fail";
		String strResult = "Announcement Portlet is not available";
	
		/*****************
		 * check Announcement header title content format
		 ********************************************************/
try {
		if (ValidationFactory.isElementPresent(wbelAnnouncementFramePath)) {
			LogFactory.info("Announcement portlet is present");

			if (GenericFactory.headerTitleFormat(AnnouncementTitleContents)) {
				strFlag = "Pass";
				strResult = "All Announcement header titles are displayed in the format --> Published Date ("+ BaseClass.dateformat + ") : Announcement Title ";
			}
			else {
				strFlag = "Fail";
				strResult = "Invalid Hearder or Date format found in the Announcement Portlet ";
			}
			
			ReportFactory.reporterOutput(strTCID,"Verify Announcement header titles on the Announcement Portlet ",  "Data format is : " + BaseClass.dateformat,
					"Announcement header title should be in the format --> Published Date ("+ BaseClass.dateformat + ") : Announcement Title ",
					strResult, strFlag);
			
		} else {
					ReportFactory.reporterOutput(strTCID,"Verify Announcement header titles on the Announcement Portlet ",  "Data format is : " + BaseClass.dateformat,
						"Announcement header title should be in the format --> Published Date ("+ BaseClass.dateformat + ") : Announcement Title", strResult,
					strFlag);
		}

	}catch (Exception e) {
		LogFactory.error("e");
		String er = e.getMessage().toString().trim();
			ReportFactory.reporterOutput(strTCID,"Verify Announcement header titles on the Announcement Portlet ", "NA","NA", er, strFlag);
		}
	}
 

	public static void verifyAnnouncementEmbededlinks(String strTCID) throws Throwable {
		
        int respCode = 200;
        List<String> emptyLinks = new ArrayList<String>();
        List<String> brokenLinks = new ArrayList<String>();
        List<String> CorrectLinks = new ArrayList<String>();
		List<WebElement> lstWebElement = GenericFactory.getLinksFromFrame(wbelAnnouncementFramePath);
		try {
			
		
        if (lstWebElement.size() > 0) {
               for (int i = 0; i < lstWebElement.size(); i++) {
                     String url = BaseClass.wbDriver.findElement(By.linkText(lstWebElement.get(i).getText()))
                                   .getAttribute("href");
                     String urlName = lstWebElement.get(i).getText();
                     if (url == null || url.isEmpty()) {
                            emptyLinks.add(urlName);
                     }
                     try {
                            HttpURLConnection huc = (HttpURLConnection) (new URL(url).openConnection());
                            huc.setRequestMethod("HEAD");
                            huc.connect();
                            respCode = huc.getResponseCode();
                            if (respCode >= 400) {
                                   brokenLinks.add(urlName);
                            } else
                                   CorrectLinks.add(urlName);
                     } catch (Exception e) {
                            e.getMessage();
                     }
               }
               System.out.println("Broken Links are :" + brokenLinks + ", Empty/null links are :" + emptyLinks
                            + ", Correct links are :" + CorrectLinks);
               if (emptyLinks.isEmpty() && brokenLinks.isEmpty()) {
                     String correctLinksString = String.join(",", CorrectLinks);
                     ReportFactory.reporterOutput(strTCID,"Verify embeded links on Announcement Portlet", "NA",
                                   "Embeded links should not be broken or empty",
                                   "Embeded links are not broken and working fine as expected :" + correctLinksString, "Pass");
               } else {
                     String brokenLinksString = String.join(",", brokenLinks);
                     String stringEmptyLinks = String.join(",", emptyLinks);
                     ReportFactory.reporterOutput(strTCID,"Verify Embeded Links on Announcement Portlet","NA",
                                   "Embeded links should not be broken or empty",
                                   "Broken Links/Empty links are :" + brokenLinksString + " , " + stringEmptyLinks, "Fail");
               }
        } else {
               System.out.println("No embeded links are present");
               
               ReportFactory.reporterOutput(strTCID,"Verify Embeded Links on Announcement Portlets.","NA",
                            "Embeded links should not be broken or empty", "No links are present", "Pass");
        }
        
		}catch (Exception e) {
			LogFactory.error("e");
			String er = e.getMessage().toString().trim();
			ReportFactory.reporterOutput(strTCID,"Verify Embeded Links on Announcement Portlets.", "NA","NA", er, "Fail");
			}
 }

	
	public static void verifyAnnouncemntsShowingInDescendingOrderOnDateTime( String strTCID,String strDepartmentNameFromExcel) throws Throwable {
		
		String strFlag = "Fail";
		String strResult = "";
		String strResultPresentOnDepartment="";
		String strResultNotPresentOnDepartment="";
		
		List<String> liMatchingDepartments=new ArrayList<String>();
		String strMatchingDepartments = String.join(",", liMatchingDepartments);
		List<String> liNonMatchingDepartments=new ArrayList<String>();
		String strNonMatchingDepartments = String.join(",", liNonMatchingDepartments);
		List<String> liDepartmentNameFromExcel = GenericFactory.splitString(strDepartmentNameFromExcel, ",");

		try {
			
			if(!strDepartmentNameFromExcel.equalsIgnoreCase("NA")) {
				
				if (liDepartmentNameFromExcel.size() > 0) {
				
					for (int j = 0; j < liDepartmentNameFromExcel.size(); j++) {
						WebElement department = GenericFactory.getDeptname(liDepartmentNameFromExcel.get(j));
						
						if (department != null) { 
							department.click();
						} else {
									GenericFactory.clickOnFlyOutDeptMenu(liDepartmentNameFromExcel.get(j));
						}
							/***********
							 * Get list of Announcement and verify if sorted according to date
							 ***************/
							if (ValidationFactory.isElementPresent(wbelAnnouncementFramePath)) {
								List<WebElement> numberOfAnnouncement = AnnouncementTitleContents;
								LogFactory.info("Number of Announcement are :" + numberOfAnnouncement.size());
								
								List<String> AnnouncementHeaderDates = GenericFactory.getListOfDatesByFrame(wbelAnnouncementFramePath,"Announcement");
								
								if (GenericFactory.verifyDateSortedOrder(AnnouncementHeaderDates) == true) {
									liMatchingDepartments.add(liDepartmentNameFromExcel.get(j));
									strResultPresentOnDepartment = strResultPresentOnDepartment.concat(liDepartmentNameFromExcel.get(j));
								} else {
									liNonMatchingDepartments.add(liDepartmentNameFromExcel.get(j));
									strResultNotPresentOnDepartment=strResultNotPresentOnDepartment.concat(liDepartmentNameFromExcel.get(j));
								}
							} 
							else {
									LogFactory.info("Announcement portlet is not present on page:");
									strResultNotPresentOnDepartment=strResultNotPresentOnDepartment.concat(liDepartmentNameFromExcel.get(j));
									liNonMatchingDepartments.add(liDepartmentNameFromExcel.get(j));
							}
					}
				
					if(liMatchingDepartments.size()>0 ){   
						strFlag = "Pass";
						strResult="Announcement title dates are displayed in descending order on active departments :"+liMatchingDepartments;
					}
					else if(liMatchingDepartments.size()==0 || liNonMatchingDepartments.size()>0){
						strFlag = "Fail";
						strResult="Announcement title dates are NOT displayed in descending order on departments  :"+liNonMatchingDepartments;
					}
					ReportFactory.reporterOutput(strTCID,"Verify Announcement displayed in descending order with newest one's on top on department ", strDepartmentNameFromExcel,"Announcements should be displayed in descending order with newest one's on top on active department pages", strResult, strFlag);
				}
			} 
			else {

					GenericFactory.navigateToHomePage();
					LogFactory.info("Verifying Announcements are  displaying in descending order on date and time in homepage");
					/*********** * Get list of Announcements and verify if sorted according to date ***************/
					List<WebElement> numberOfAnnouncement = BaseClass.wbDriver.findElements(By.xpath("//div[@class='section warning']"));
					List<String> AnnouncementHeaderDates = GenericFactory.getListOfDatesByFrame(wbelAnnouncementFramePath,"Announcement" );
		
					if (GenericFactory.verifyDateSortedOrder(AnnouncementHeaderDates) == true) {
						
						strFlag = "Pass";
						strResult = "Announcement title dates are in descending order with newest one's on top by date on homepage";
					} 
					else {
							LogFactory.info("Announcement title dates are not in sorted order on homepage");
							strResult = "Announcement title dates are not in sorted order on homepage";
					}
					ReportFactory.reporterOutput(
							strTCID,"Verify Announcements are displayed in descending order with newest one's on top of homepage", 
							"NA", "Announcements should be displayed in descending order with newest one's on top of homepage",
							strResult, strFlag);
				}
			} catch (Throwable e) {
				ReportFactory.reporterOutput(strTCID,"Verify Announcements are displayed in descending order with newest one's on top of homepage", "NA", "NA",
					e.getMessage().toString(), "Fail");
			}

	}
	public static void verifyAnnouncementTitleOnDepartmentUncheck(String wcmDeptName, String titleName, String strTCID,
            String wcmCopyToDepartment) throws Throwable {
     
     List<String> listOfDepartments = new ArrayList<>();
     if (!wcmDeptName.equals("General")) {
            listOfDepartments.add(wcmDeptName);
     }
     if (!wcmCopyToDepartment.equals("NA")) {
            listOfDepartments.addAll(GenericFactory.splitString(wcmCopyToDepartment, ","));
     }
     String strResult = "Fail";
     String commaSeprateListOfDepartments=String.join(", ", listOfDepartments);
     String actualResult = "Announcement title is still displayed " + titleName + " after unchecking the "
                  +commaSeprateListOfDepartments + " department";
     
     if (wcmDeptName.equalsIgnoreCase("General") && wcmCopyToDepartment.equalsIgnoreCase("NA")) {
    	 actualResult = "Unable to verify this scenario, since the department is genereal and also it's not copied to any of the deparments.";
            strResult = "Pass";
     } else {
            
            if (verifyAnnouncementHeader(titleName)) {
            	GenericFactory.checkDepartmentMyPreference();
                  GenericFactory.uncheckDepartmentMyPreference(listOfDepartments);
                  List<String> listActualData = new ArrayList<String>();
                  for (int i = 0; i < PortalLeftNavigation_POF.ListAllActiveLinks.size(); i++) {
                         String temp = PortalLeftNavigation_POF.ListAllActiveLinks.get(i).getText();
                         listActualData.add(temp);
                  }
                  if (!listOfDepartments.contains(listActualData)) {
                         if (verifyAnnouncementHeader(titleName)) {
                                strResult = "Fail";
                                actualResult = "<B>Announcement title : </B>" + titleName + "is present even after the <B>"
                                              + commaSeprateListOfDepartments+ "</B>  departments are filtered ";
                                System.out.println("Fail");
                         } else {
                                strResult = "Pass";
                                actualResult = "<B>Announcement title : </B>" + titleName + "is not present after the <B>"
                                              +commaSeprateListOfDepartments+ "  </B>departments are filtered ";
                                System.out.println("Pass");
                         }
                  } else {
                         actualResult = "Department <B>" + commaSeprateListOfDepartments
                                       + " </B>are not disabled after unchecking them due to some error.";
                         strResult = "Fail";

                  }
            } else {
                  actualResult = "Announcement is supposed to be present on announcement portlet but it's not present";
                  strResult = "Fail";

            }
     }
     ReportFactory.reporterOutput(strTCID,
                  "To verify announcement on the announcement portlet when the respective department is unchecked",
                  "<B>Title name is :</B></br>" + titleName
                  +"</br><B>WCM Department name is :</B></br>" + wcmDeptName
                  +"</br><B>WCM Copy to Department name is :</B></br>" + wcmCopyToDepartment,
                  "Announcements should not be displayed after un-checking relevant department", actualResult,
                  strResult);
}
	public static void verifyAnnouncementHeaderCount(String strTCID) throws Throwable {
        String flagResult = "Pass";
        String actualResult = "Announcement header is not present.";
        try {
               
        for(int i=0; i<HeaderList.size(); i++) {
               if(HeaderList.get(i).getText().contains("(")) {
                     String temp=GenericFactory.splitString(HeaderList.get(i).getText(), "(").get(1);
                     String temp1=GenericFactory.splitString(temp, ")").get(0);
                     int headerCountFromHeader = Integer.parseInt(temp1);
                     int actualHeaderCountFromTitleList=AnnouncementTableTitle.size();
                     if(actualHeaderCountFromTitleList==headerCountFromHeader) {
                            flagResult = "Pass";
                            actualResult = "Announcement header count is :<B>"+headerCountFromHeader+"</B>"
                                         + "</br>Actual count of Announcement title is :<B>"+actualHeaderCountFromTitleList+"</B>" ;
                            }
                     else {
                            flagResult = "Fail";
                            actualResult = "Announcement header count is :<B>"+headerCountFromHeader+"</B></br>"
                                         + "</br>Actual count of Announcement present is :<B>"+actualHeaderCountFromTitleList+"</B>" ;
                         
                     }
               }
        }
        ReportFactory.reporterOutput(strTCID, "Verify Announcement header count is same as announouncements available", "NA",
                     "Announcement header count should be same as announcement present", actualResult, flagResult);
 
        }
        catch (Exception e) {
               ReportFactory.reporterOutput(strTCID, "Verify Announcement header count is same as announouncements available", "NA",
                            "NA", "Unable to verify this scenario due to below error</br>"+e.getMessage(), "Fail");
        
        }
        
       
         }
	 /**
     * This method verifies that title of relevant Announcement should not visible
     * if user un-check his product type from product filter icon.
     * 
      * @author shrey.choudhary
     * @createdAt 28-06-2018
     * @param UserProductName
     * @param WcmProducts
     * @param titleName
     * @return boolean
     * @throws Throwable
     */
	public static void verifyAnnouncementTitleOnProductUncheck(String UserProductName, String WcmProducts,
            String titleName, String strTCID) throws Throwable {
     String strAnnouncementFlag = "Fail";
     String strResult = "<p style=\"color:ROYALBLUE;\">Announcement title :</p>" + titleName
                  + "is still present after the corresponding product types are filtered " + WcmProducts;
     List<String> listUserProducts = GenericFactory.splitString(UserProductName, ",");
     List<String> listWcmProducts = GenericFactory.splitString(WcmProducts, ",");
     
     GenericFactory.checkAllProducts();
     if (verifyAnnouncementHeader(titleName)) {

            LogFactory.info("Announcement link is present");
            ValidationFactory.getElementIfPresent(By.xpath("//div[@id='js-segments']")).click();
            for (int i = 0; i < listUserProducts.size(); i++) {
                  List<String> listUserProductTemp = GenericFactory
                                .strUserProductToMatchWithWCMList(listUserProducts.get(i));
                  if (GenericFactory.verifyIfProductMatch(listWcmProducts, listUserProductTemp)) {
                         GenericFactory.productUncheck((GenericFactory.getParentProductSegmnent(listUserProducts.get(i))));
                  }
            }
            ApplyFilterButton.click();
            if (!verifyAnnouncementHeader(titleName)) {
                  strAnnouncementFlag = "Pass";
                  strResult = "<B>Announcement title :</B>" + titleName + "is not present after the corresponding " + String.join(", ", WcmProducts)
                                + " product types are filtered";
            }
            ReportFactory.reporterOutput(strTCID,
                         "Verify Announcement title should not be visible to the user, if user unchecked related products tagged with WCM content",
                         "<B>User products : </B>"+String.join(", ", listUserProducts)+""
                         +"<B>WCM products :<B>"+String.join(", ", WcmProducts)
                         ,
                         "Announcements should not be displayed after unchecking " + String.join(", ", WcmProducts) + " product types",
                         strResult, strAnnouncementFlag);
     } else {
            LogFactory.info("Unable to find the test data to execute the test case" + titleName);
            strAnnouncementFlag = "Pass";
            strResult = "<B>Announcement title :</B>" + titleName
                         + " is not visible to the user as given in test data sheet :";
            ReportFactory.reporterOutput(strTCID,
                         "Verify Announcement title should not be visible to the user, if user unchecked related products tagged with WCM content",
                         "<B>User products : </B>"+String.join(", ", listUserProducts)
                                       +"<B>WCM products : </B>"+String.join(", ", WcmProducts),
                         "Announcements should not be displayed after unchecking :<B></br>" + String.join(", ", WcmProducts) + "</B> product types",
                         strResult, strAnnouncementFlag);
     }
}

 

	public static boolean checkReadMoreAndCollapseLinkForAnnouncement(String strTCID) throws Throwable {
		String strFlag = "Fail";
		String strResult = "Readmore link is not present";
		boolean booFlagReadmore = true;
		boolean booFlagCollapse = true;
		boolean readMoreLinkflag = true;

		try {
			List<WebElement> readMoreLink = wbelAnnouncementFramePath
					.findElements(By.xpath(".//div[@class='secondary-action-container']"));
			int readmorelinkcount = readMoreLink.size();

			if (readmorelinkcount > 0) {
				for (int j = 0; j < readmorelinkcount; j++) {
					Thread.sleep(2000);
					readMoreLink.get(j).click();
					Thread.sleep(2000);
					String strCollapseText = readMoreLink.get(j).getText().trim().toString();
					booFlagReadmore = true;

					if (strCollapseText.equalsIgnoreCase("Collapse")) {
						booFlagCollapse = true;
						Thread.sleep(2000);
						readMoreLink.get(j).click();
						Thread.sleep(2000);
						String strReadmoreText = readMoreLink.get(j).getText().trim().toString();
					
						if (strReadmoreText.equalsIgnoreCase("Read More")) {
							booFlagReadmore = true;
						} else {
							booFlagReadmore = false;
						}
						break;
					} else {
							booFlagReadmore = true;
							booFlagCollapse = false;
							break;
					}
				}
			}
			else {
				booFlagReadmore = false;
			}

			if (booFlagReadmore && booFlagCollapse) {
				strFlag = "Pass";
				strResult = "ReadMore link is displayed for Announcement portlet and working as expected";
			} else if (booFlagReadmore && !booFlagCollapse) {
				strFlag = "Fail";
				strResult = "ReadMore link is present for Announcement portlet but not working as expected";
			} else {
				strFlag = "Pass";
				strResult = "ReadMore link is not displayed for Announcement portlet";
			}
			ReportFactory.reporterOutput(strTCID,
					"Verify 'ReadMore & Collapse' links are displayed on Announcement portlet", "NA",
					"If 'ReadMore & Collapse' link is present then it should work as expected", strResult, strFlag);

		} catch (Exception e) {
				ReportFactory.reporterOutput(strTCID,
					"Verify 'ReadMore & Collapse' links are displayed on Announcement portlet", "NA","NA", e.getMessage(), "Fail");
		}
		return booFlagCollapse && booFlagReadmore;

	}
    
 

    public static void verifyExpandAndCollapseLinkOnAnnuoncementPortlet(String strTCID) throws Throwable {

        String strFlag = "Pass";
        String strResult = "Announcements count are less then 6, hence not verifying this scenario.";

        boolean booFlagExpand = false;
        boolean booFlagCollapse = true;

        try {
               int listOfWebElements = AnnouncementTableTitle.size();
               if (listOfWebElements > 6) {
                     if (ValidationFactory.isElementPresent(linkAnnouncementOfExpand)) {
                            linkAnnouncementOfExpand.click();

                            if (linkAnnouncementOfExpand.getText().equalsIgnoreCase("Collapse")) {
                                   LogFactory.info("Collapse link is present");
                                   booFlagCollapse = true;
                                   Thread.sleep(2000);
                                   linkAnnouncementOfExpand.click();
                                   Thread.sleep(2000);
                                   if (linkAnnouncementOfExpand.getText().equalsIgnoreCase("Expand")) {
                                          booFlagExpand = true;
                                   } 
                            } else {
                                   LogFactory.info("Collapse link is not present");
                                   booFlagCollapse = false;
                            }
                     }
               } else {
                     LogFactory.info("Expand link is not displayed");

               }
               if (booFlagExpand && booFlagCollapse) {
                     LogFactory.info("Expand link  and collapse link are  present");
                     strFlag = "Pass";
                     strResult = "Expand link is displayed for announcement portlet and working as expected";
               } else if (booFlagExpand && !booFlagCollapse) {

                     LogFactory.info("Expand link is displayed but collapse link is not present");
                     strFlag = "Fail";
                     strResult = "Expand link is present for announcement portlet but not working as expected";
               }
               ReportFactory.reporterOutput(strTCID, "Verify Expand link displayed for announcement portlet on my dealerpath",
                            "NA", "If 'Expand' link is present it should work as expected", strResult, strFlag);
        }

        catch (Exception e) {
               ReportFactory.reporterOutput(strTCID, "Verify Expand link displayed for announcement portlet on my dealerpath",
                            "NA", "If 'Expand' link is present it should work as expected", "Unable to verify this scenario due to this error :"+ e.getMessage(), "Fail");
        
        }


}
    
}   