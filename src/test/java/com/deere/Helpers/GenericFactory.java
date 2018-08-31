
/* 
 * Project    : DealerPath
 * Script     : GenericFactory
 * Author     : Shrishail Baddi
 * Date       : April.08.2018
 * Last Modified On:
 * Modified By :
 */

package com.deere.Helpers;

import java.io.File;
import java.io.IOException;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.List;
import java.util.StringTokenizer;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;
import org.testng.Assert;

import com.deere.PageFactory.Alerts_POF;
import com.deere.PageFactory.Announcements_POF;
/*import com.deere.PageFactory.Alerts_POF;
import com.deere.PageFactory.Announcements_POF;*/
import com.deere.PageFactory.PortalLeftNavigation_POF;

public class GenericFactory extends BaseClass{

	/*
	 * A String is truncated according to reg and stored in a List, and finally
	 * returns a List
	 *
	 * @param str
	 * 
	 * @param reg
	 * 
	 * @return List<String>
	 * 
	 */

	// public static final String DEFAULT_REG = ",";

	public static List<String> splitString(String str, String reg) {

		List<String> strList = new ArrayList<String>();

		if (isNull(str)) {
			return strList;
		}

		StringTokenizer st = new StringTokenizer(str, reg);

		while (st.hasMoreElements()) {

			strList.add(st.nextToken().trim());
		}

		return strList;

	}

	/*
	 * method to determine if a string is null or 0 Yes returns true otherwise false
	 * 
	 * @param str
	 * 
	 * @return boolean
	 */

	public static boolean isNull(String str) {
		if (str == null || str.length() < 1) {
			return true;
		}

		return false;
	}

	// public static WebElement element=null;

	@SuppressWarnings("null")
	public static void impersonateUser(String strRACFID) {

		WebDriver driver = BaseClass.wbDriver;

		try {

			if (!isNull(strRACFID)) {

				if (ValidationFactory.isElementPresent(By.xpath("//b[text() = 'Admin Links']"))) {

					// System.out.println(ValidationFactory.getElementIfPresent(By.xpath("//div[@id='Impersonate_User'
					// and @class ='active']")));

					ValidationFactory.getElementIfPresent(By.xpath("//a[text() = 'Impersonate User']")).click();
					LogFactory.info("Clicked Impersonate User link from left navigation.....");
					WaitFactory.waitForPageLoaded();

					LogFactory.info("Enter User RACF ID.....");
					ValidationFactory.getElementIfPresent(By.xpath("//input [@name = 'inputText_AU']"))
							.sendKeys(strRACFID);

					LogFactory.info("Click on Search Button.....");
					ValidationFactory.getElementIfPresent(By.xpath("//input [ @name = 'Search']")).click();

					if (ValidationFactory.isElementPresent(By.xpath("//input[@type='radio' and @name='userDNS']"))) {

						LogFactory.info("Select the user.....");
						ValidationFactory.getElementIfPresent(By.xpath("//input[@type='radio' and @name='userDNS']"))
								.click();

						if (ValidationFactory
								.validateButtonEnable(By.xpath("//input[@type='button' and @name='Impersonate']"))) {

							ValidationFactory
									.getElementIfPresent(By.xpath("//input[@type='button' and @name='Impersonate']"))
									.click();
							LogFactory.info("User " + strRACFID + " successfully logged into appllication ");
						} else {
							LogFactory.error("Unable to locate the button Impersonate/ it is disabled");
						}
					} else {
						LogFactory.error(strRACFID + " does not exist/No user Found");
					}
				}

			}

		} catch (Exception e) {
			System.out.println(e.getMessage());
		}

	}

	public static void endImpersonateOrLogoutUser() {

		WebElement userelement = null;

		try {

			userelement = ValidationFactory
					.getElementIfPresent(By.xpath("//div[contains (@class,'user-info') and @id ='js-user-info']"));

			if (userelement != null) {

				userelement.click();

				WebElement btnelement = ValidationFactory
						.getElementIfPresent(By.xpath("//button[@class='btn' and @onclick='endImpersonate(this)']"));

				if (btnelement != null) {

					btnelement.click();
				} else {
					ValidationFactory.getElementIfPresent(
							By.xpath("//button[@class='btn' and @onclick='logoutDealerPath(this)']")).click();
				}

			}

		} catch (Exception e) {
			e.printStackTrace();
			LogFactory.info("Unable to logout or impersonate the user");
			Assert.fail("Unable to logout or impersonate the user");

		}

	}

	public static void utilityMenuAdminClick() {

		try {

			WebElement userelement = null;

			userelement = ValidationFactory
					.getElementIfPresent(By.xpath("//div[contains (@class,'user-info') and @id ='js-user-info']"));

			if (userelement != null) {

				userelement.click();
				LogFactory.info("Utility Menu Clicked.....");

			//	WaitFactory.waitForElementClickable(ValidationFactory.getElementIfPresent(By.xpath("//a[contains (@href ,'dealerpath.admin')]")));
			
				ValidationFactory.getElementIfPresent(By.xpath("//a[contains (@href ,'dealerpath.admin')]")).click();
				LogFactory.info("Successfully clicked on Admin Link from utility manu....");

			}

		} catch (Exception e) {
			e.printStackTrace();
			LogFactory.info("Unable to click on Admin utility link or Element not found");
			Assert.fail("Unable to click on Admin utility link or Element not found");

		}
	}

	public static void utilityMenuMyPreferenceClick() {

		try {

			WebElement userelement = null;

			userelement = ValidationFactory
					.getElementIfPresent(By.xpath("//div[contains (@class,'user-info') and @id ='js-user-info']"));

			if (userelement != null) {

				userelement.click();
				LogFactory.info("Utility Menu Clicked.....");

				ValidationFactory.getElementIfPresent(By.xpath("//a[@id='preferences']")).click();
				LogFactory.info("Successfully clicked on My Preference Link from utility manu....");

			}

		} catch (Exception e) {
			e.getMessage();
			LogFactory.info("Unable to click on My Preference utility link or Element not found");
			Assert.fail("Unable to click on My Preference utility link or Element not found");

		}

	}

	public static void utilityMenuSwitchSiteClick() {

		try {

			WebElement userelement = null;

			userelement = ValidationFactory
					.getElementIfPresent(By.xpath("//div[contains (@class,'user-info') and @id ='js-user-info']"));

			if (userelement != null) {

				userelement.click();
				LogFactory.info("Utility Menu Clicked.....");

				ValidationFactory.getElementIfPresent(By.xpath("//a[ @id='switch']")).click();
				LogFactory.info("Successfully clicked on switch site Link from utility manu....");

			}

		} catch (Exception e) {
			e.printStackTrace();
			LogFactory.info("Unable to click on switch site  utility link or Element not found");
			Assert.fail("Unable to click on switch site utility link or Element not found");

		}

	}

	public static void getUserPrefredLanguage() {

		try {

			WebElement element = null;

			utilityMenuMyPreferenceClick();

			element = ValidationFactory.getElementIfPresent(By.xpath("//select[@id='lang']"));

			if (element != null) {

				Select preferedLang = new Select(element);

				BaseClass.strUserPrefLang = preferedLang.getFirstSelectedOption().getText();
				ValidationFactory.getElementIfPresent(By.xpath("//button[@id='preference-cancel']")).click();
				LogFactory.info("User current prefered  language is....." + BaseClass.strUserPrefLang);

			} else {

				BaseClass.strUserPrefLang = "default";
				LogFactory.info(
						"User do not have the option to change the  prefered language, so site is refereing to the default laguage as per the site Region.....");
				ValidationFactory.getElementIfPresent(By.xpath("//button[@id='preference-cancel']")).click();
			}

		} catch (Exception e) {
			e.getMessage();
			LogFactory.info("Unable to click on switch site  utility link or Element not found");
			// Assert.fail("Unable to click on switch site utility link or Element not
			// found");

		}

	}

	/**
	 * Method to get the list of Dates by header name for Alert and announcement
	 * 
	 * @author shrey.choudhary
	 * @param header
	 * @createdAt 26-05-2018
	 * @return
	 */
	/*
	 * public static List<String> getListOfDatesByHeaderName(String header) {
	 * 
	 * 
	 * 
	 * String headerAlert = Alerts_POF.wbelAlertFramePath.findElement(By.xpath(
	 * "//div[@class='section-header']/h3")).getText(); String headerAnnouncement =
	 * Announcements_POF.wbelAnnouncementFramePath.findElement(By.xpath(
	 * "//div[@class='section-header']/h3")) .getText(); List<String> dateList = new
	 * ArrayList<>(); if (headerAlert.equalsIgnoreCase(header)) { List<WebElement>
	 * elementsList = Alerts_POF.wbelAlertFramePath.findElements(By.xpath(
	 * ".//*[@class='list-item-title']")); for (int i = 0; i < elementsList.size();
	 * i++) { String temp = elementsList.get(i).getText();
	 * dateList.add(GenericFactory.splitString(temp, ":").get(0)); } return
	 * dateList; } else if (headerAnnouncement.equalsIgnoreCase(header)) {
	 * List<WebElement> elementsList = Announcements_POF.wbelAnnouncementFramePath
	 * .findElements(By.xpath(".//*[@class='list-item-title']")); for (int i = 0; i
	 * < elementsList.size(); i++) { String temp = elementsList.get(i).getText();
	 * dateList.add(GenericFactory.splitString(temp, ":").get(0)); } return
	 * dateList; } else {
	 * LogFactory.info("Unable to perform the sort order on dates."); } return null;
	 * }
	 * 
	 */
	/**
	 * Method to verify the format of dates on the Alert & Announcement portlet
	 * 
	 * @author shrey.choudhary
	 * @param header
	 * @createdAt 24-05-2018
	 * @return
	 */
	public static boolean isValidDateFormat(String dateformat, List<String> listOfDates) {
		boolean isValidFormat = false;
		for (int i = 0; i < listOfDates.size(); i++) {
			Date date = null;
			try {
				SimpleDateFormat sdf = new SimpleDateFormat(dateformat);
				date = sdf.parse(listOfDates.get(i));
				if (!listOfDates.get(i).equals(sdf.format(date))) {
					date = null;
				}
			} catch (ParseException ex) {
				ex.printStackTrace();

			}
			if (date == null) {
				isValidFormat = false;
				break;
			} else {
				isValidFormat = true;
			}
		}
		return isValidFormat;
	}

	/**
	 * This Method verifies the order of dates on the Alert & Announcement portlet
	 * 
	 * @author shrey.choudhary
	 * @param header
	 * @createdAt 27-05-2018
	 * @return
	 */
	public static boolean verifyDateSortedOrder(List<String> listOfDatesInFrame) {
		List<String> ExpectedlistOfDates = new ArrayList<>();
		
		ExpectedlistOfDates.addAll(listOfDatesInFrame);
		
		Collections.sort(ExpectedlistOfDates, Collections.reverseOrder());

		System.out.println("Actual list is : ");
		listOfDatesInFrame.forEach(System.out::println);
		System.out.println("sorted list is : ");
		ExpectedlistOfDates.forEach(System.out::println);
		if (listOfDatesInFrame.equals(ExpectedlistOfDates)) {
			return true;
		}
		return false;
	}

	/**
	 * This method to find the list of webElements from a frame by xpath
	 * 
	 * @author shrey.choudhary
	 * @param WebElement
	 * @createdAt 27-05-2018
	 * @return List<WebElement>
	 */

	public static List<WebElement> getLinksFromFrame(WebElement framePath) {
		List<WebElement> getWebElementsLinks = framePath.findElements(By.tagName("a"));
		List<WebElement> listOfElements = new ArrayList<WebElement>();
		for (int i = 0; i < getWebElementsLinks.size(); i++) {
			listOfElements.add(getWebElementsLinks.get(i));
		}
		return listOfElements;
	}

	public static void createHeaderSection(String header) throws InterruptedException {

		ReportFactory.addRoottable();
		ReportFactory.reportSectionName(header);
		ReportFactory.reporterOutputHeader();

	}

/*	public static void navigateToHomePage() throws InterruptedException {

		WebElement userelement = null;
		userelement = ValidationFactory.getElementIfPresent(By.xpath(".//*[@class='app-title']"));

		if (userelement != null) {
			LogFactory.info("Navigating to the DealerPath Homepage");
			userelement.click();
		}

	}
	
	*/
	
	public static void navigateToHomePage() throws InterruptedException {
		  WebElement userelement = null;
		  if ( ValidationFactory.getElementIfPresent(By.xpath("//input[@class='filter-box fav-filter']"))!= null ) {
			  LogFactory.info("User is on homepage");
		  } else {
		   userelement = ValidationFactory.getElementIfPresent(By.xpath(".//*[@class='app-title']"));
		   if (userelement != null) {
		    LogFactory.info("Navigating to the DealerPath Homepage");
		    userelement.click();
		   }
		  }
		 }

	public static List<String> getCheckBoxValues() {
		// List<WebElement> getWebElementsLinks =
		// BaseClass.wbDv.findElements(By.xpath(".//*[@class='checkbox-value']/div/label/div"));

		List<WebElement> getWebElementsLinks = BaseClass.wbDriver
				.findElements(By.xpath("//form[@id='productSegmentsForm']/div/div[@class='value']//div"));
		// System.out.println(getWebElementsLinks.size());
		// System.out.println(getWebElementsLinks);
		List<String> listOfElements = new ArrayList<String>();
		for (int i = 0; i < getWebElementsLinks.size(); i++) {

			// System.out.println(">>>>>>>>>>>>>>"+getWebElementsLinks.get(i).getText());
			listOfElements.add(getWebElementsLinks.get(i).getText());

		}
		// System.out.println(listOfElements);

		return listOfElements;
	}

	/**
	 * This method is to capture the screenshot and save them in configured
	 * directory, i.e confi.property or user.dir
	 * 
	 * @author shrey.choudhary
	 * @createdAt 24-05-2018
	 * @param TCID
	 * @throws Exception
	 * 
	 * @ModifiedBy
	 * @ModifiedAt
	 */
	public static String getScreenshot(String TCID) throws Exception {
		String filePath = "";

		if (BaseClass.ENABLE_SCREENSHOT.equalsIgnoreCase("ON")) {
			BaseClass.strScreenshotDir = FileFolderFactory.createDirectory(BaseClass.strScreenshotDir,
					"\\Screenshot\\");

			DateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy_HH-mm-ss");
			Date date = new Date();

			File scrFile = ((TakesScreenshot) BaseClass.wbDriver).getScreenshotAs(OutputType.FILE);
			try {
				// now copy the screenshot to desired location using copyFile //method
				filePath = BaseClass.strScreenshotDir + TCID + "_" + dateFormat.format(date) + ".png";
				FileUtils.copyFile(scrFile, new File(filePath));
			} catch (IOException e) {
				System.out.println(e.getMessage());
			}
		}
		return filePath;

	}

	/*
	 * public static String productlistcomparer(String userDefinedProducts, String
	 * wcmProducts) {
	 * 
	 * String flag = "Fail"; try {
	 * 
	 * String str1 =
	 * "Agriculture/Ag Equipment,Agriculture/Sprayers & Nutrient Applicators,Construction/Construction,Forestry/Forestry,Hitachi/Mining,Hitachi/Hitachi"
	 * ;
	 * 
	 * String str2 = "Agriculture/Ag Equipment,Construction,Hitachi";
	 * 
	 * 
	 * List<String> userDefinedProductsList =
	 * Arrays.asList(userDefinedProducts.split("\\s*,\\s*")); List<String>
	 * wcmProductsList = Arrays.asList(wcmProducts.split("\\s*,\\s*"));
	 * 
	 * 
	 * for (int i = 0; i < wcmProductsList.size(); i++) {
	 * 
	 * String wcmProductsvalue = wcmProductsList.get(i);
	 * 
	 * for(int k=0;k<userDefinedProductsList.size();k++) {
	 * 
	 * String userDefinedProductsvalue = userDefinedProductsList.get(k);
	 * 
	 * if(userDefinedProductsvalue.contains(wcmProductsvalue)) { flag = "Pass";
	 * System.out.println(wcmProductsList.get(i) + " Is Same as " +
	 * userDefinedProductsList.get(k));
	 * LogFactory.info("Product details matched Successfully."); break;
	 * 
	 * 
	 * }
	 * 
	 * }
	 * 
	 * }
	 * 
	 * for (int i = 0; i < userDefinedProductsList.size(); i++) {
	 * 
	 * String userDefinedProductsvalue = userDefinedProductsList.get(i);
	 * 
	 * for (int j = 0; j < wcmProductsList.size(); j++) {
	 * 
	 * String wcmProductsvalue = wcmProductsList.get(j);
	 * 
	 * if (wcmProductsvalue.contains(userDefinedProductsvalue.toString()) ||
	 * userDefinedProductsvalue.contains(wcmProductsvalue.toString())) { flag =
	 * "Pass"; System.out.println (wcmProductsvalue + " Is Same as " +
	 * userDefinedProductsvalue);
	 * LogFactory.info("Product details matched Successfully."); break;
	 * 
	 * }
	 * 
	 * }
	 * 
	 * } return flag;
	 * 
	 * }
	 * 
	 * catch (Exception e) { flag = "Fail"; LogFactory.info("Fail"); //
	 * Assert.fail("Fail");
	 * 
	 * } return flag;
	 * 
	 * }
	 */

/*	public static boolean productlistcomparer(String userDefinedProducts, String wcmProducts) {

		boolean flag = false;
		try {
			
			 * String str1 =
			 * "Agriculture/Ag Equipment,Agriculture/Sprayers & Nutrient Applicators,Construction/Construction,Forestry/Forestry,Hitachi/Mining,Hitachi/Hitachi"
			 * ;
			 * 
			 * String str2 = "Agriculture/Ag Equipment,Construction,Hitachi";
			 

			LogFactory.info("Verifying whether user products are matching with WCM product types");

			List<String> userDefinedProductsList = Arrays.asList(userDefinedProducts.split("\\s*,\\s*"));
			List<String> wcmProductsList = Arrays.asList(wcmProducts.split("\\s*,\\s*"));

			
			 * for (int i = 0; i < wcmProductsList.size(); i++) {
			 * 
			 * String wcmProductsvalue = wcmProductsList.get(i);
			 * 
			 * for(int k=0;k<userDefinedProductsList.size();k++) {
			 * 
			 * String userDefinedProductsvalue = userDefinedProductsList.get(k);
			 * 
			 * if(userDefinedProductsvalue.contains(wcmProductsvalue)) { flag = "Pass";
			 * System.out.println(wcmProductsList.get(i) + " Is Same as " +
			 * userDefinedProductsList.get(k));
			 * LogFactory.info("Product details matched Successfully."); break;
			 * 
			 * 
			 * }
			 * 
			 * }
			 * 
			 * }
			 

			for (int i = 0; i < userDefinedProductsList.size(); i++) {

				String userDefinedProductsvalue = userDefinedProductsList.get(i);

				for (int j = 0; j < wcmProductsList.size(); j++) {

					String wcmProductsvalue = wcmProductsList.get(j);

					if (wcmProductsvalue.contains(userDefinedProductsvalue.toString())
							|| userDefinedProductsvalue.contains(wcmProductsvalue.toString())) {
						flag = true;
						// System.out.println(wcmProductsvalue + " Is Same as " +
						// userDefinedProductsvalue);

						break;

					}

				}

			}

			String isCountryMatched = flag ? "Matched" : "Not Matched";
			LogFactory.info(" User products are " + isCountryMatched + " with WCM product type");

			return flag;

		}

		catch (Exception e) {
			flag = false;

			String isCountryMatched = flag ? "Matched" : "Not Matched";
			LogFactory.info(" User country is " + isCountryMatched + " with WCM country");

			return flag;

		}

	}
*/
	public static WebElement getDeptname(String deptname) {

		// ********* List for active links *****************************
		List<WebElement> deptListActiveLinks = new ArrayList<WebElement>();
		WebElement strdepnameinactive = null;
		WebElement strdepnameactive = null;

		for (int i = 0; i < PortalLeftNavigation_POF.ListAllActiveLinks.size(); i++) {
			strdepnameactive = PortalLeftNavigation_POF.ListAllActiveLinks.get(i);
			deptListActiveLinks.add(strdepnameactive);
			// System.out.println(deptListActiveLinks.get(i).getText());
		}

		// ********** List for inactive links ******************************
		List<WebElement> deptListInactiveLinks = new ArrayList<WebElement>();
		for (int i = 0; i < PortalLeftNavigation_POF.listAllInactiveLinks.size(); i++) {
			strdepnameinactive = PortalLeftNavigation_POF.listAllInactiveLinks.get(i);
			deptListInactiveLinks.add(strdepnameinactive);
			// System.out.println(strdepnameinactive.getText());
		}

		// *************** compare excel data with active and inactive
		// links*****************

		boolean flag1 = false;
		boolean flag2 = false;

		LogFactory.info(" Vearifying the department " + deptname + " is active or Not...");

		int j = 0;
		for (; j < deptListActiveLinks.size(); j++) {
			// System.out.println(deptList.get(j).getText() + "::" + testdata);

			if (deptListActiveLinks.get(j).getText().equals(deptname)) {
				flag1 = true;
				break;
			}

		}

		if (flag1 != true) {
			for (int k = 0; k < deptListInactiveLinks.size(); k++) {
				// System.out.println(deptListInactiveLinks.size());
				if (deptListInactiveLinks.get(k).getText().equals(deptname)) {
					flag2 = true;
					break;
				}

			}
		}

		if (flag1 == true && flag2 == false) {
			// LogFactory.info(" Department "+ deptname + " is active ...");
			return deptListActiveLinks.get(j);
		} else if (flag2 == true && flag1 == false) {
			// LogFactory.info(" Department "+ deptname + " is not active ...");
			return null;
		} else {
			// LogFactory.info(" Department "+ deptname + " is not active ...");
			return null;
		}
	}

	/*
	 * public static void toClickonDept(String deptName) { try{
	 * 
	 * // To clik on 'My DealerPath' arrow WebElement arrowClick =
	 * BaseClass.wbDriver.findElement(By.xpath(".//*[@id='left_nav_0']"));
	 * 
	 * 
	 * //Mouse hover Actions Actions obj = new Actions(BaseClass.wbDriver);
	 * obj.moveToElement(arrowClick).build().perform();
	 * 
	 * 
	 * 
	 * //To get list of all departments List<WebElement> deptNameList =
	 * BaseClass.wbDriver.findElements(By.xpath(".//*[@class='flyout']/li/a"));
	 * for(int i=1; i< deptNameList.size(); i++) { String deptNameListValue =
	 * deptNameList.get(i).getText().toString().trim(); //
	 * System.out.println(deptNameListValue);
	 * LogFactory.info("dept Name List Values= " +deptNameListValue);
	 * 
	 * //To click on desire department name given
	 * if(deptNameListValue.contains(deptName.trim())) {
	 * 
	 * deptNameList.get(i).click();
	 * LogFactory.info("Clicked on desired deparment name. " +deptName); break; } }
	 * } catch(Exception e) { LogFactory.info("Click on the department. "+e); } }
	 */

	public static void clickOnFlyOutDeptMenu(String UserDeptName) {
		try {

			String SelectedLeftTab = BaseClass.wbDriver
					.findElement(By.xpath("//*[@class='active  selected leftNavDeptPadding']")).getText();

			if (!UserDeptName.equals(SelectedLeftTab)) {
				Actions builder = new Actions(BaseClass.wbDriver);
				WebElement arrowClick = BaseClass.wbDriver.findElement(By.xpath(".//*[@id='left_nav_0']"));
				builder.moveToElement(arrowClick).build().perform();

				List<WebElement> actualDeptNameList = BaseClass.wbDriver
						.findElements(By.xpath(".//*[@class='flyout']/li/a"));
				for (int i = 1; i < actualDeptNameList.size(); i++) {
					String deptNameListValue = actualDeptNameList.get(i).getText().toString().trim();
					if (deptNameListValue.contains(UserDeptName.trim())) {

						actualDeptNameList.get(i).click();
						LogFactory.info(" Clicking on the active department " + UserDeptName + " from Portlet links page");
						break;
					}
				}
			}
		} catch (Exception e) {
			LogFactory.info("Clicking on the department. " + e);
		}
	}

	/*public static boolean countryListComparison(String UserCountry, String wcmCountry) {

		boolean flag = false;

		List<String> wcmCountryList = Arrays.asList(wcmCountry.split("(,)"));
		// System.out.println(wcmCountryList);

		List<String> userCountryList = Arrays.asList(UserCountry.split("(,)"));
		// System.out.println(userCountryList);

		try {

			LogFactory.info("Verifying user country is matching with WCM country");

			for (int i = 0; i < wcmCountryList.size(); i++) {
				for (int j = 0; j < userCountryList.size(); j++) {

					// System.out.println(wcmCountryList.get(j).concat("/")+"::"+userCountryList.get(i).split("/").trim());
					if (!wcmCountryList.get(j).contains("/")) {
						String arr[] = userCountryList.get(j).split("/");
						if (wcmCountryList.get(i).contains(arr[0])) {
							flag = true;
							// LogFactory.info("User country contains MRU country");

						} else {

							// LogFactory.info("User country doesnot contain WCM country");

						}
					} else if (!userCountryList.get(j).contains("/")) {
						if (wcmCountryList.get(i).contains("/")) {
							flag = false;
							// LogFactory.info("User country doesnot contains MRU country");
						}
					}

					else {
						// System.out.println(wcmCountryList.get(i) + "::" + userCountryList.get(j));
						if (wcmCountryList.get(i).contains(userCountryList.get(j).trim())) {
							flag = true;
							// LogFactory.info("User country contains WCM country");

						} else {

							// LogFactory.info("User country doesnot contain WCM country");

						}

					}
				}
			}
		} catch (Exception e) {
			LogFactory.info("Exception message :-->" + e.getMessage());
			flag = false;
		}

		String isCountryMatched = flag ? "Matched" : "Not Matched";
		LogFactory.info(" User country is " + isCountryMatched + " with WCM country");

		return flag;
	}
*/
	
	
	
	/**
	 * Method to get the list of Dates by header name for Alert and announcement
	 * Author: shrey.choudhary
	 * 
	 * @param header
	 * @return
	 */
	public static List<String> getListOfDatesByHeaderName(String header) {

		String headerAlert = Alerts_POF.wbelAlertFramePath.findElement(By.xpath("//div[@class='section-header']/h3"))
				.getText();
		
		String headerAnnouncement = Announcements_POF.wbelAnnouncementFramePath
				.findElement(By.xpath("//div[@class='section-header']/h3")).getText();
		List<String> dateList = new ArrayList<>();
		if (headerAlert.equalsIgnoreCase(header)) {
			List<WebElement> elementsList = Alerts_POF.wbelAlertFramePath
					.findElements(By.xpath(".//*[@class='list-item-title']"));
			for (int i = 0; i < elementsList.size(); i++) {
				String temp = elementsList.get(i).getText();
				dateList.add(GenericFactory.splitString(temp, ":").get(0));
			}
			return dateList;
		} else if (headerAnnouncement.equalsIgnoreCase(header)) {
			List<WebElement> elementsList = Announcements_POF.wbelAnnouncementFramePath
					.findElements(By.xpath(".//*[@class='list-item-title']"));
			for (int i = 0; i < elementsList.size(); i++) {
				String temp = elementsList.get(i).getText();
				dateList.add(GenericFactory.splitString(temp, ":").get(0));
			}
			return dateList;
		} else {
			LogFactory.info("Unable to perform the sort order on dates.");
		}
		return null;
	}

	
	
	public static List<String> getListOfDatesByFrame(WebElement alertannouncement, String portletName) {

		
/*		String headerAlert = Alerts_POF.wbelAlertFramePath.findElement(By.xpath("//div[@class='section-header']/h3")).getText();
		String headerAnnouncement = Announcements_POF.wbelAnnouncementFramePath.findElement(By.xpath("//div[@class='section-header']/h3")).getText();
*/		
		List<String> dateList = new ArrayList<>();
		
		if ((ValidationFactory.isElementPresent(alertannouncement)) && (portletName == "Alert")) {
					
					List<WebElement> elementsList = Alerts_POF.wbelAlertFramePath.findElements(By.xpath(".//*[@class='list-item-title']"));
					for (int i = 0; i < elementsList.size(); i++) {
							String temp = elementsList.get(i).getText();
							dateList.add(GenericFactory.splitString(temp, ":").get(0));
					}
					return dateList;
		}
		else  {		List<WebElement> elementsList = Announcements_POF.wbelAnnouncementFramePath.findElements(By.xpath(".//*[@class='list-item-title']"));
					for (int i = 0; i < elementsList.size(); i++) {
						String temp = elementsList.get(i).getText();
						dateList.add(GenericFactory.splitString(temp, ":").get(0));
					}
					return dateList;
					
			  } 

	}
	
	
	public static void setUserPreferredLang() throws InterruptedException {

		String strExcelValue = "English";

		if (!strExcelValue.isEmpty()) {

			utilityMenuMyPreferenceClick();

			Thread.sleep(1000);

			WebElement element = ValidationFactory.getElementIfPresent(By.xpath("//select[@id='lang']"));
			if (element != null) {
				Select preferedLang = new Select(element);
				BaseClass.strUserPrefLang = preferedLang.getFirstSelectedOption().getText();

				if (!BaseClass.strUserPrefLang.equals(strExcelValue)) {

					preferedLang.selectByVisibleText(strExcelValue);
					ValidationFactory.getElementIfPresent(By.xpath("//button[@id='preference-save']")).click();

					LogFactory.info("Language from excel sheet " + strExcelValue + " is selected");
				} else {

					LogFactory.info("Excel value and selected value on My Preferences is same");

					ValidationFactory.getElementIfPresent(By.xpath("//button[@id='preference-cancel']")).click();
				}
			} else {
				LogFactory.info(
						"User do not have the option to change the  prefered language, so site is refereing to the default laguage as per the site Region.....");
				ValidationFactory.getElementIfPresent(By.xpath("//button[@id='preference-cancel']")).click();
			}
		} else {
			LogFactory.info("Current language on My Preferences is English");

		}
	}

	
	
	public static boolean userAndWCMCountryComparison(String UserCountry, String wcmCountry) {
        boolean flag = false;
        List<String> wcmCountryList = Arrays.asList(wcmCountry.split("(,)"));
        try {
               for (int i = 0; i < wcmCountryList.size(); i++) {
                     if (!wcmCountryList.get(i).contains("/")) {
                            String arr[] = UserCountry.split("/");
                            if (wcmCountryList.get(i).contains(arr[0].toString().trim())) {
                                   flag = true;
                            }
                     } else if (!UserCountry.contains("/")) {
                            if (wcmCountryList.get(i).contains(UserCountry.toString().trim())) {
                                   flag = true;
                            }
                     } else if (UserCountry.contains("/")) {
                            int Usercounter = 0;
                            for (int j = 0; j < UserCountry.length(); j++) {
                                   if (UserCountry.charAt(j) == '/') {
                                          Usercounter++;
                                   }
                            }
                            int WCMcounter = 0;
                            for (int j = 0; j < wcmCountryList.get(i).length(); j++) {
                                   if (wcmCountryList.get(i).charAt(j) == '/') {
                                          WCMcounter++;
                                   }
                            }
                            if (Usercounter == 1 && WCMcounter == 1) {
                                   String arr[] = UserCountry.split("/");
                                   if (wcmCountryList.get(i).contains(arr[1].toString().trim())) {
                                          flag = true;
                                   }
                            } else if (Usercounter == 1 && WCMcounter == 2) {
                                   String arr[] = UserCountry.split("/");
                                   if (wcmCountryList.get(i).contains(arr[1].toString().trim())) {
                                          flag = true;
                                   }
                            } else if (Usercounter == 2 && WCMcounter == 2) {
                                   String arr[] = UserCountry.split("/");
                                   if ((wcmCountryList.get(i).contains(arr[0].toString().trim())) && (wcmCountryList.get(i).contains(arr[1].toString().trim()))
                                                 && (wcmCountryList.get(i).contains(arr[2].toString().trim()))) {
                                          flag = true;
                                   }
                            } else if (Usercounter == 2 && WCMcounter == 1) {
                                   String arr[] = UserCountry.split("/");
                                   if (wcmCountryList.get(i).contains(arr[1].toString().trim())) {
                                          flag = true;
                                   }
                            }
                     } else {
                            if (wcmCountryList.get(i).contains(UserCountry.trim().toString().trim())) {
                                   flag = true;
                            }
                     }
               }
        } catch (Exception e) {
               LogFactory.info("Exception message :-->" + e.getMessage());
               flag = false;
        }
        return flag;
 }

 @SuppressWarnings("null")
 public static List<String> strUserProductToMatchWithWCMList(String parentProductUser) {
        List<String> childProductAfterSplit = new ArrayList<>();
        // for (int i = 0; i < parentProductUser.size(); i++) {
        if (parentProductUser.contains("/")) {
               childProductAfterSplit.add(GenericFactory.splitString(parentProductUser, "/").get(0));
               childProductAfterSplit.add(GenericFactory.splitString(parentProductUser, "/").get(1));
        } else {
               childProductAfterSplit.add(parentProductUser);
               childProductAfterSplit.add("null");

        }
        return childProductAfterSplit;
 }

 public static boolean verifyIfProductMatch(List<String> parentProductWCM, List<String> childProductAfterSplit) {
        for (int i = 0; i < parentProductWCM.size(); i++) {
               if (parentProductWCM.get(i).contains("/")) {
                     String parentProductAfterSplitWCM = GenericFactory.splitString(parentProductWCM.get(i), "/").get(0);
                     String childProductAfterSplitWCM = GenericFactory.splitString(parentProductWCM.get(i), "/").get(1);
                     if ((childProductAfterSplit.get(0).equalsIgnoreCase(parentProductAfterSplitWCM))
                                   && ((childProductAfterSplitWCM.equalsIgnoreCase(childProductAfterSplit.get(1)))
                                                 || ((childProductAfterSplit.get(1).equals("null"))))) {

                            return true;
                     }

               } else if (parentProductWCM.get(i).equalsIgnoreCase(childProductAfterSplit.get(0))) {
                     return true;
               }
        }
        return false;
 }

 public static boolean userAndWCMProductTypeComparison(String UserProduct, String wcmProduct) {
        List<String> listOfParentProductUser = GenericFactory.splitString(UserProduct, ",");
        List<String> listOfWcmProduct = GenericFactory.splitString(wcmProduct, ",");
        for (int i = 0; i < listOfParentProductUser.size(); i++) {
               List<String> ParentProductUser = strUserProductToMatchWithWCMList(listOfParentProductUser.get(i));
               if (verifyIfProductMatch(listOfWcmProduct, ParentProductUser)) {
                     return true;
               }
        }
        return false;
 }


 /**
	 * This method is to parent product from the given product segments
	 * @author Neeraja.mantri
	 * @createdAt 22-06-2018
	 * @param product
	 * @return
	 * @throws Throwable
	 * @ModifiedBy
	 * @ModifiedAt
	 */
	
		public static String getParentProductSegmnent(String product)
	{
			String productName=product;
		if(product.contains("/")) {
			List<String> temp = splitString(product, "/");
			productName=temp.get(1);
		}
			
		 if (productName.matches("(?i)Construction|Utility"))	
		{
			
			return "Construction";
		}
			
		else if(productName.matches("(?i)Forestry"))	
		{
			return "Forestry";
			
		}
		else if(productName.matches("(?i)Homeowner"))	
		{
			return "Homeowner";
			
		}
		else if(productName.matches("(?i)CWP"))	
		{
			return "CWP";
			
		}	
			
		else if(productName.matches("(?i)Hitachi|Mining"))	
		{
			return "Hitachi";
			
		}	
		
		else if(productName.matches("(?i)Agriculture|Ag Equipment|Sprayers & Nutrient Applicators|Scraper and Scraper Tractor|Forage Harvester"))	
		
			return "Agriculture";
			
		
		
		else if(productName.matches("(?i)Commercial|Residential"))	
		{
			return "Turf";
			
		}	
		else if(productName.matches("(?i)Golf"))	
		{
			return "Golf";
			
		}
		
		return null;
	}	
	

		/**
		 * This method is to parent product from the given product segments
		 * @author Neeraja.mantri
		 * @createdAt 22-06-2018
		 * @param product
		 * @return
		 * @throws Throwable
		 * @ModifiedBy
		 * @ModifiedAt
		 */
	
	public static void productUncheck(String product)
	{
		
		//ProductSegment_POF.wbelProductSegmentIcon.click();
		
		List<String> listOfProductSegments=getCheckBoxValues();

		
		for(int i=0;i<listOfProductSegments.size();i++)
		{
				WebElement products=ValidationFactory.getElementIfPresent(By.xpath(".//*[@id='productSegmentsForm']/div/div[" + (i + 2) + "]"));
				String temp = ValidationFactory.getElementIfPresent(By.xpath(".//*[@id='productSegmentsForm']/div/div[" + (i + 2) + "]")).getText().toString().trim();
				WebElement getCheckboxPath = BaseClass.wbDriver.findElement(By.xpath(".//*[@id='"+temp+"']"));
				if(product.equalsIgnoreCase(listOfProductSegments.get(i)) &&  getCheckboxPath.isSelected()){
					products.click();	
					break;
				}
		}
			
	}
	
	/**
	 * This method is to check the alerts/announcements title with format
	 * @author Neeraja.mantri
	 * @createdAt 21-06-2018
	 * @param by
	 * @return
	 * @throws Throwable
	 * @ModifiedBy
	 * @ModifiedAt
	 */
	

public static boolean headerTitleFormat(List<WebElement> title) throws Throwable {

		boolean dateflag = false;
		boolean titleflag = false;
		
		//List<WebElement> title = wbDriver.findElements(by);
		List<String> headertitles = new ArrayList<String>();

		for (int i = 0; i < title.size(); i++) {

			headertitles.add(title.get(i).getText());
			String strText = title.get(i).getText();
			
			if (title.get(i).getText().contains(":")) {
				
				String rowValues[] = title.get(i).getText().split(":");
				LogFactory.info("Date :" + rowValues[0].trim());
				String dateToValidate = rowValues[0].trim();

				if (dateToValidate != null) {
					
					if (DateFactory.isThisDateValid(dateToValidate, BaseClass.dateformat)) {
						dateflag = true;
						LogFactory.info("Date format is valid");
					
					} else {
							dateflag = false;
							LogFactory.info("Date format is not valid");
							break;
					}		
				}

				if (rowValues[1].length() > 0) {
					LogFactory.info("-Title---------" + rowValues[1].trim());
					titleflag = true;
				}
				else {
						titleflag = false;
						break;
					}

				// LogFactory.info("Content is in required format");
			} else {

				dateflag=false;
				break;
			}
		
		}

		return dateflag && titleflag;

	}
	
/*	public static boolean headerTitleFormat(By by) throws Throwable {

		boolean flag = false;

		List<WebElement> title = BaseClass.wbDriver.findElements(by);

		List<String> headertitles = new ArrayList<String>();

		for (int i = 0; i < title.size(); i++) {

			// System.out.println("Alerts title list" +
			// alertsTitle.get(i).getText().toString());
			headertitles.add(title.get(i).getText());

			String strText = title.get(i).getText();

			if (title.get(i).getText().contains(":")) {

				String rowValues[] = title.get(i).getText().split(":");
				// System.out.println(alertsTitle.size()+">>>>>>>>>>>>>>"+rowValues[0].trim());
				LogFactory.info("Dates :" + rowValues[0].trim());
				String dateToValidate = rowValues[0].trim();
				flag = true;
				if (dateToValidate != null) {
					boolean validDate = DateFactory.isThisDateValid(dateToValidate, BaseClass.dateformat);
					flag = true;
					LogFactory.info("validDate=" + validDate);
				} else {

					LogFactory.info("Date format is not valid");
				}

				if (rowValues[1].length() > 0) {
					LogFactory.info("-Titles---------" + rowValues[1].trim());
					flag = true;
				}

				// LogFactory.info("Content is in required format");
			} else {

				LogFactory.info("Content is not in the required format");
			}
		}

		return flag;

	}
*/



	public static void checkAllProducts() {
		  ValidationFactory.getElementIfPresent(By.xpath("//div[@id='js-segments']")).click();
		  List<String> listOfProductSegments = getCheckBoxValues();
		  for (int i = 0; i < listOfProductSegments.size(); i++) {
		   WebElement products = ValidationFactory
		     .getElementIfPresent(By.xpath(".//*[@id='productSegmentsForm']/div/div[" + (i + 2) + "]"));
		   String temp = ValidationFactory
		     .getElementIfPresent(By.xpath(".//*[@id='productSegmentsForm']/div/div[" + (i + 2) + "]")).getText()
		     .toString().trim();
		   WebElement getCheckboxPath = BaseClass.wbDriver.findElement(By.xpath(".//*[@id='" + temp + "']"));
		   if (!getCheckboxPath.isSelected()) {
		    products.click();
		   }
		  }
		  alertPageFactory.ApplyFilterButton.click();
		 }    
 
  
	/**
	 * This method is to get active second level department link
	 * @author Neeraja.mantri
	 * @createdAt 28-06-2018
	 * @param strTestdata
	 * @return WebElement
	 * @ModifiedBy
	 * @ModifiedAt
	 */
	
	public static WebElement  getActive2ndLevelDepartment(String strTestdata)
	{
		List<WebElement> secondLevelActiveDepartment = new ArrayList<WebElement>();
		int i = 0;
		for (; i < alertPageFactory.SecondLevelDepartment.size(); i++) 
		{
			WebElement strdepartmentnameactive = alertPageFactory.SecondLevelDepartment.get(i);
			secondLevelActiveDepartment.add(strdepartmentnameactive);
		
		
		if(alertPageFactory.SecondLevelDepartment.get(i).getText().equals(strTestdata))
		{
			LogFactory.info("Second level department matching is :" + alertPageFactory.SecondLevelDepartment.get(i).getText());
			break;
		}
		else {
			
			LogFactory.info("Second level department is not matching :" + alertPageFactory.SecondLevelDepartment.get(i).getText());
			
		}
		}
		
		return alertPageFactory.SecondLevelDepartment.get(i);
		
	}
 
	/**
	 * This method will click on a particular index page after verifying the department matched
	 * @author Archana.gaikwad
	 * @createdAt 02-07-2018
	 * @param strSecondlevelDepartment
	 * @param strIndexPage
	 * @ModifiedBy Neeraja.mantri
	 */
	
	public static boolean clickOnIndexPageLinkOnSecondLevelDepartment(String strSecondlevelDepartment,String strIndexPage) {
		
		
		
		List<WebElement> liActualSubDepartmentsFrame = BaseClass.wbDriver.findElements(By.xpath(".//*[@id='links-target']/div"));
		   String strHeaderNameFlag="Fail";
		   String strResult = "Fail";
		   boolean returnFlag=false;
		    
		   for (int j = 0; j < liActualSubDepartmentsFrame.size(); j++) 
		   {
		    String strSubDepartment = liActualSubDepartmentsFrame.get(j).getText().trim();
		    String[] lines = strSubDepartment.split("\n");
		    String strheadername = lines[0];
		    //System.out.println(strheadername);
		    WebElement weSecondLevelLinks = liActualSubDepartmentsFrame.get(j);
		   // System.out.println(liActualSubDepartmentsFrame.get(j).getText());
		    
		    	if (strSecondlevelDepartment.equals(strheadername)) 
		    	{
		    	strHeaderNameFlag= "Pass";
		      
		    	List<WebElement> links = weSecondLevelLinks.findElements(By.tagName("a"));
		      
		    	for (WebElement we : links)
		    	{
		    		if (we.getText().equals(strIndexPage))
		    			{
		    			LogFactory.info("Index page link is present in the subdepartment :"+strSecondlevelDepartment);
		    			WebElement weThirdLevelIndexLink = BaseClass.wbDriver.findElement(By.linkText(strIndexPage));
		    			weThirdLevelIndexLink.click();
		    			returnFlag=true;
		    			break;
		    			}
		    		else {
		    			
		    			LogFactory.info("Index Page link is not present in the subdepartment :"+strSecondlevelDepartment);
		    			
		    		}
		    		
		       }
		      break;  
		      }
		    	else {
		    		
		    		LogFactory.info("Secondlevel department name is not matching with expected");
		    	}
		      
		       break;
		      }
		   
		   return returnFlag;
	
	}
	    

	public boolean uncheckProductList(String wcmProductList) throws Throwable {

		boolean flag=true;
		BaseClass.wbDriver.findElement(By.xpath("//div[@id='js-segments']")).click();

		List<WebElement> productType = BaseClass.wbDriver
				.findElements(By.xpath("//div[@id='js-segments-popover']//div[@class='value']"));

		List<String> excel = new ArrayList<String>();
		excel.add("Agriculture");
		excel.add("CWP");
		excel.add("Construction");
		excel.add("Forestry");
		excel.add("Hitachi");
		excel.add("Turf");

		List<WebElement> checkBox = BaseClass.wbDriver.findElements(
				By.xpath("//div[@id='js-segments-popover']//div[@class='value']//input[@type='checkbox']"));

		for (int i = 0; i < checkBox.size(); i++) {
			if (!checkBox.get(i).isSelected()) {
				System.out.println("*********unchecked**********");
				checkBox.get(i).click();
			}
		}

		for (int j = 0; j < excel.size(); j++) {
			for (int i = 0; i < productType.size(); i++) {
				if (excel.get(j).equalsIgnoreCase(productType.get(i).getText())) {
					checkBox.get(i).click();
				}
			}
		}

		BaseClass.wbDriver.findElement(By.xpath(".//*[@id='js-segments-popover']/div[2]/div[3]/button")).click();

		String errorMessage = BaseClass.wbDriver
				.findElement(By.xpath("//div[@class='popover-content']//p[@id='productSegmentsError']")).getText();

		System.out.println("errorMessage===" + errorMessage);
		if(errorMessage!=null){
			flag= false;
			System.out.println("flag"+flag);
		
		}
		return flag;
}
	
	public static boolean verifyDealerType(String flagDealerType, String wcmDealerType) {
		  if(flagDealerType.equalsIgnoreCase(wcmDealerType)) {
		   return true;
		  }
		  else if (flagDealerType.contains("/")||wcmDealerType.contains("/") || wcmDealerType.equalsIgnoreCase("NA")) {
		   return true;
		  }
		  return false;
		 }	
	
	
	public static boolean clickOnDepartmentByName(String DepartmentName) {
		try {
			
			
			for (int i = 0; i < PortalLeftNavigation_POF.ListAllActiveLinks.size(); i++) {
				if (PortalLeftNavigation_POF.ListAllActiveLinks.get(i).getText().equalsIgnoreCase(DepartmentName)) {
					PortalLeftNavigation_POF.ListAllActiveLinks.get(i).click();
					return true;
				}
			}
			System.out.println("Given department name : " + DepartmentName
					+ " is not available/disabled in Left navigation window");
		} catch (Exception e) {
			LogFactory
					.info("Given department name : " + DepartmentName + " is not available in Left navigation window");
		}
		return false;
	}
	public static void uncheckDepartmentMyPreference(List<String> departmentName) throws InterruptedException {
        GenericFactory.utilityMenuMyPreferenceClick();
        Thread.sleep(5000);
        for (int i = 0; i < announcementFactory.allCheck.size(); i++) {
               if (!(announcementFactory.allCheck.get(i).getAttribute("checked") == null)
                            && departmentName.contains(announcementFactory.allCheck1.get(i).getText())) {
            	   announcementFactory.allCheck.get(i).click();
                    
               }
        }
        Thread.sleep(5000);
        BaseClass.wbDriver.findElement(By.xpath("//*[@id='preference-save']")).click();
        Thread.sleep(15000);
 }
	
 public static void checkDepartmentMyPreference() throws InterruptedException {
        GenericFactory.utilityMenuMyPreferenceClick();
        Thread.sleep(5000);
        for (int i = 0; i < announcementFactory.allCheck.size(); i++) {
               if (announcementFactory.allCheck.get(i).getAttribute("checked") == null){
            	   announcementFactory.allCheck.get(i).click();
               }
        }
 }
 
 public static boolean toClickonDeptOnFavourite(String deptName)
 {
     boolean Flag = false;
    BaseClass.wbDriver.findElement(By.xpath(".//*[@id='main-header']/div[1]/div/h1")).click();
      
   //To get list of all departments
   List<WebElement> deptNameList = BaseClass.wbDriver.findElements(By.xpath(".//*[@id='leftNav']/li/a"));
   for(int i=0; i<=deptNameList.size()-1; i++)
   {
    String deptNameListValue = deptNameList.get(i).getText().toString().trim();
   
    if(deptNameListValue.equalsIgnoreCase(deptName))
    {
        if(deptNameList.get(i).isEnabled())
        {
        	deptNameList.get(i).click();
        	LogFactory.info("Clicked on desired deparment name. " +deptName);
        
        	Flag= true;
      
        }  
        break;
   }
}
return Flag; 
  
}

 
 
/* public static boolean toClickonDeptOnFavourite(String deptName)
 {
     boolean flag = false;
    BaseClass.wbDriver.findElement(By.xpath(".//*[@id='main-header']/div[1]/div/h1")).click();
      
   //To get list of all departments
   List<WebElement> deptNameList = BaseClass.wbDriver.findElements(By.xpath(".//*[@id='leftNav']/li/a"));
   for(int i=0; i<=deptNameList.size()-1; i++)
   {
    String deptNameListValue = deptNameList.get(i).getText().toString().trim();
   
    if(deptNameListValue.equalsIgnoreCase(deptName))
    {
        if(deptNameList.get(i).isEnabled())
        {
        	deptNameList.get(i).click();
        	LogFactory.info("Clicked on desired deparment name. " +deptName);
        
        	flag= true;
      
        }  
        break;
   }
}
return flag; 
  
}*/
	
}
