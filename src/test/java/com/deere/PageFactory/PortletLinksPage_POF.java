/* 
 * Project    : DealerPath
 * Script     : Homepage_POF
 * Author     : Neeraja Mantri
 * Date       : May.15.2018
 * Last Modified On:
 * Modified By :
 */

package com.deere.PageFactory;

import java.net.HttpURLConnection;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.How;
import org.testng.Assert;
import org.testng.asserts.SoftAssert;
import java.util.Map.Entry;

import com.deere.Helpers.BaseClass;
import com.deere.Helpers.GenericFactory;
import com.deere.Helpers.LogFactory;
import com.deere.Helpers.ReportFactory;
import com.deere.Helpers.ValidationFactory;
import com.deere.Helpers.WaitFactory;

/**
 * @author neeraja.mantri
 *
 */
public class PortletLinksPage_POF {

	static WebDriver HomDriver;
	static SoftAssert softAssert = new SoftAssert();

	public PortletLinksPage_POF(WebDriver driver) {
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

	/**
	 * @param title
	 * @param TCID
	 * @throws Throwable
	 */
	public static void checkUserLogIntoHomepage(String strTitle, String strTCID) throws Throwable {

		String strTitlePage = HomDriver.getTitle();

		WaitFactory.waitForPageLoaded();

		LogFactory.info(" Title " + strTitlePage + " is displayed on the homepage");
		String strFlag = "FAIL";
		String strResult = "Title is not displayed";

		try {
			if (ValidationFactory.isElementPresent(wbelTitleOfHomePage)) {
				if (strTitlePage.equalsIgnoreCase(strTitle)) {

					strFlag = "PASS";
					strResult = "Title is displayed :" + strTitle;
				}

			}
			ReportFactory.reporterOutput(strTCID, "Verify the title is displayed on the homepage", strTitle,
					"Title on homepage should be  " + strTitlePage, strResult, strFlag);

	/*		if (strFlag.equalsIgnoreCase("FAIL")) {
				Assert.assertFalse(true);
			}*/

		} catch (Throwable e) {

			ReportFactory.reporterOutput(strTCID, "Verify the title is displayed on the homepage", strTitle,
					"Verify the title is displayed on the homepage " + strTitlePage, e.getMessage().toString(),
					strFlag);
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
	public static void getDealerPrincipalRole(String strTestData, String strTCID) throws Throwable {

		String strFlag = "Fail";
		String strResult = "Dealer is not having dealer principal role";
		try {
			if (strTestData.equals("NA")) {
				ReportFactory.reporterOutput(strTCID, "Verify Dealer Principal role.", "NA",
						"Verify Dealer Principal role.", strResult, strFlag);
			} else {

				List<String> frameList = new ArrayList<String>();
				for (int i = 0; i < PortalLeftNavigation_POF.ListAllActiveLinks.size(); i++) {
					String temp = PortalLeftNavigation_POF.ListAllActiveLinks.get(i).getText();
					frameList.add(temp);
				}
				if (frameList.contains(strTestData)) {
					strFlag = "Pass";
					strResult = "Dealer is having dealer principal role";
				}
				ReportFactory.reporterOutput(strTCID, "Verify Dealer Principal role.", "NA",
						"Verify Dealer Principal role.", strResult, strFlag);
			}

		} catch (Throwable e) {

			ReportFactory.reporterOutput(strTCID, "Verify Dealer Principal role.", "NA",
					"Verify Dealer Principal role.", e.getMessage().toString(), strFlag);
		}
	}

	/*
	 * public static void PortletLinks(String UserDefinedCountry, String WCMCountry,
	 * String UserDefinedProducts, String WCMProducts, String RootsiteArea, String
	 * Contenttype, String DepartmentName, String level2, String Title) throws
	 * Throwable { String flagPT = "";
	 * 
	 * 
	 * 
	 * boolean titleFlag=false; boolean sectionFlag=false;
	 * 
	 * String testCaseDescription = " Verify link portlet for content type : " +
	 * Contenttype; String inputValue = " <B>WCM Maping :</B> Country : " +
	 * WCMCountry + " ,Product Type : " + WCMProducts + " ,Dept : " + DepartmentName
	 * + " ,Sub Category :  " + level2 +"<P>" + " <B>Dealer Maping :</B> Country : "
	 * + UserDefinedCountry + " ,Product Type : " + UserDefinedProducts;
	 * 
	 * try {
	 * 
	 * LogFactory.info("Verify Country Check."); if
	 * (UserDefinedCountry.contains(WCMCountry)) {
	 * 
	 * LogFactory.info("Country details Matched Successfully.");
	 * 
	 * if (RootsiteArea.contains("AT_Portletlinks")) {
	 * 
	 * LogFactory.info("Site Area Matched Successfully.");
	 * 
	 * // *************************This is Product //
	 * Check*****************************
	 * 
	 * LogFactory.info("Verify Product Check");
	 * 
	 * //flagPT = GenericFactory.productlistcomparer(UserDefinedProducts,
	 * WCMProducts);
	 * 
	 * if (flagPT.equalsIgnoreCase("Pass")) { // *************************This is
	 * Department // Check*****************************
	 * 
	 * LogFactory.info("Verify Department Name");
	 * 
	 * WebElement department = GenericFactory.getDeptname(DepartmentName);
	 * 
	 * if (department != null)
	 * 
	 * department.click();
	 * 
	 * else GenericFactory.toClickonDept(DepartmentName);
	 * 
	 * List<WebElement> actualheadername1 = HomDriver
	 * .findElements(By.xpath(".//*[@id='links-target']/div"));
	 * 
	 * for (int j = 0; j < actualheadername1.size(); j++) { String temp1 =
	 * actualheadername1.get(j).getText().trim(); String[] lines =
	 * temp1.split("\n"); String headername = lines[0]; WebElement secondLevel =
	 * actualheadername1.get(j);
	 * 
	 * if (level2.equals(headername)) {
	 * 
	 * // System.out.println(level2); sectionFlag=true; List<WebElement> links =
	 * secondLevel.findElements(By.tagName("a"));
	 * 
	 * for (WebElement we : links) { //System.out.println("links>>>>>>>>>>" +
	 * we.getText());
	 * 
	 * if (Title == null || Title == "") { LogFactory.info("Title is blank.");
	 * ReportFactory.reporterOutput("Portlet Links", "Verify Portlet Links.:Title.",
	 * Title, "Title should be present.", "Title is missing.", "Fail"); break; }
	 * 
	 * if (we.getText().equals(Title)) { titleFlag=true; // WebElement link2 =
	 * secondLevel. // System.out.println(we.getAttribute("href")); String
	 * cextension = we.getAttribute("href");
	 * 
	 * String s = cextension.substring(cextension.lastIndexOf("/") + 1); //
	 * System.out.println(s); String s1 =
	 * cextension.substring(cextension.lastIndexOf(".") + 1); //
	 * System.out.println(s1); String s2 = s1.substring(0, 4); //
	 * System.out.println(s2);
	 * 
	 * switch (Contenttype) { case "AT-Link": if (we.getAttribute("href") != null) {
	 * int respCode = 200; HttpURLConnection huc = (HttpURLConnection) (new
	 * URL(cextension) .openConnection()); huc.setRequestMethod("HEAD");
	 * huc.connect(); respCode = huc.getResponseCode();
	 * 
	 * if (respCode <= 400) {
	 * 
	 * LogFactory.info("This is AT-Link.");
	 * ReportFactory.reporterOutput("Portlet Links",testCaseDescription,
	 * inputValue,Title, we.getText(), "Pass"); } else {
	 * LogFactory.info("This is AT-Link.");
	 * ReportFactory.reporterOutput("Portlet Links",testCaseDescription,
	 * inputValue,Title, "It's a Broken link and response code is : " + respCode,
	 * "Fail"); } break; }
	 * 
	 * 
	 * else {
	 * 
	 * LogFactory.info("This is not AT-Link.");
	 * ReportFactory.reporterOutput("Portlet Links",testCaseDescription, inputValue,
	 * Title, we.getText(), "Fail");
	 * 
	 * }
	 * 
	 * break; case "AT-Document": if (s2.contains("exe") || s2.contains("doc") ||
	 * s2.contains("pdf") || s2.contains("bmp") || s2.contains("jpg") ||
	 * s2.contains("jpeg") || s2.contains("html")|| s2.contains("png")) {
	 * LogFactory.info("This is AT-Document.");
	 * 
	 * ReportFactory.reporterOutput("Portlet Links",testCaseDescription, inputValue,
	 * "Content Type should be AT-Document : " + s2, Title +"\n" + cextension,
	 * "Pass");
	 * 
	 * }
	 * 
	 * else { LogFactory.info("This is not AT-Document.");
	 * 
	 * ReportFactory.reporterOutput("Portlet Links",testCaseDescription, inputValue,
	 * "Content Type is AT-Document but:  " + s2 +
	 * " not found in the extension list ", cextension, "Fail");
	 * 
	 * } break;
	 * 
	 * case "AT-Rich Text":
	 * 
	 * if (s2.contains("rtf")) {
	 * 
	 * LogFactory.info("This is AT-RichText.");
	 * ReportFactory.reporterOutput("Portlet Links",testCaseDescription, inputValue,
	 * "Content Type should be AT-RichText : " + s2, cextension, "Pass"); } else {
	 * LogFactory.info("This is not AT-RichText.");
	 * ReportFactory.reporterOutput("Portlet Links",testCaseDescription, inputValue,
	 * "Content Type should be AT-RichText : " + s2,
	 * "Content Type is not AT-RichText  as .rtf is not found in the href ",
	 * "Fail");
	 * 
	 * }
	 * 
	 * break; case "Index page": LogFactory.info("Index page");
	 * ReportFactory.reporterOutput("Portlet Links",testCaseDescription, inputValue,
	 * Title , Title + cextension , "Pass"); break;
	 * 
	 * default: LogFactory.info("This is AT-Link.");
	 * ReportFactory.reporterOutput("Portlet Links",testCaseDescription, inputValue,
	 * "Content Type " + s2 , "Content Type is" + cextension, "Pass"); break; }
	 * 
	 * // break; } } if (titleFlag != true) {LogFactory.info(Title +
	 * " : Link not available in the department : " + DepartmentName +
	 * " and section : " + level2);
	 * ReportFactory.reporterOutput("Portlet Links",testCaseDescription,
	 * inputValue,Title + " : Link should be available in the department : " +
	 * DepartmentName + " and section " + level2, "Link : " + Title +
	 * " is not available in section : " + level2, "Fail");}
	 * 
	 * 
	 * }
	 * 
	 * } if (sectionFlag != true) {LogFactory.info(level2 +
	 * " : section is not available in the department : " + DepartmentName );
	 * ReportFactory.reporterOutput("Portlet Links",testCaseDescription,
	 * inputValue,level2 + " : section should be available in the department : " +
	 * DepartmentName, "Section " + level2 + " is not available in " +
	 * DepartmentName, "Fail");}
	 * 
	 * //break; }else { LogFactory.info("Products not matching.");
	 * ReportFactory.reporterOutput("Portlet Links",testCaseDescription, inputValue,
	 * "WCM Products and User Products should match.","Products didn't Match.",
	 * "Fail");
	 * 
	 * }
	 * 
	 * }
	 * 
	 * } else { LogFactory.info("Country Details are not matching.");
	 * ReportFactory.reporterOutput("Portlet Links",testCaseDescription,
	 * inputValue,"WCMCountry and UserCountry should match",
	 * "Countries didn't Match.", "Fail"); }
	 * 
	 * } catch (Exception e) {
	 * 
	 * 
	 * ReportFactory.reporterOutput("Portlet Links",testCaseDescription ,
	 * inputValue,"NA",e.getMessage(), "Fail");
	 * 
	 * } }
	 */

	public static void PortletLinks(String strWCMTCID, String userDefinedCountry, String wcmCountry,
			String userDefinedProducts, String wcmProducts, String rootSiteArea, String contentType,
			String departmentName, String secondLevel, String title) throws Throwable {

		String flagPT = "";

		boolean booCountry = GenericFactory.userAndWCMCountryComparison(userDefinedCountry, wcmCountry);
		boolean booProduct = GenericFactory.userAndWCMProductTypeComparison(userDefinedProducts, wcmProducts);

		LogFactory.info(">>>>>>>>Executing " + strWCMTCID  + " >>>>>>>>>>>" );
		LogFactory.info("Verifying the links on the Portlet links page...");

		boolean booTitleFlag = false;
		boolean sectionFlag = false;

		String testCaseDescription = " Verify link portlet for content type : " + contentType;
		String inputValue = " <B>WCM Maping :</B> Country : " + wcmCountry + " ,Product Type : " + wcmProducts
				+ " ,Dept : " + departmentName + " ,Sub Category :  " + secondLevel + "<P>"
				+ " <B>Dealer Maping :</B> Country : " + userDefinedCountry + " ,Product Type : " + userDefinedProducts;
		String expectedValue = "Alert portlet showing the valid alerts title for the user ";

		try {

			WebElement department = GenericFactory.getDeptname(departmentName);
			if (department != null) {
				LogFactory.info(" Clicking on the active department " + departmentName + " from homepage");
				department.click();
			} else { GenericFactory.clickOnFlyOutDeptMenu(departmentName);	}
			
			List<WebElement> actualHeaderName = HomDriver.findElements(By.xpath(".//*[@id='links-target']/div"));
		
			LogFactory.info("Looking for the header " + secondLevel);
			
			for (int j = 0; j < actualHeaderName.size(); j++) {
				
				String tempValue = actualHeaderName.get(j).getText().trim();
				String[] lines = tempValue.split("\n");
				String headerName = lines[0];

				WebElement secondLevelLink = actualHeaderName.get(j);

				if (secondLevel.equals(headerName)) {
				
					sectionFlag = true;
					LogFactory.info("Section/category " + secondLevel + " found in the department " + departmentName);
					//secondLevelLink.click();
			
					List<WebElement> links = secondLevelLink.findElements(By.tagName("a"));

					LogFactory.info("Looking for the link " + title + " in the category " + secondLevel );
					
					for (WebElement linkValues : links) {

						if (linkValues.getText().equals(title)) {
							booTitleFlag = true;
							LogFactory.info("Link " + title + " found in the category " + secondLevel );

							if (booCountry == true && booProduct == true && booTitleFlag == true) {

								String strLink = linkValues.getAttribute("href").toString().trim();

								String linkCharacters = strLink.substring(strLink.lastIndexOf("/") + 1);

								String linkExtention = strLink.substring(strLink.lastIndexOf(".") + 1);

								String linkExtentionType = linkExtention.substring(0, 4);

								switch (contentType) {

								case "AT-Link":
									if (strLink != null) {
										
										LogFactory.info("Content Type for the link " + title + " is AT-Link.");
										
										int respCode = 200;
										HttpURLConnection huc = (HttpURLConnection) (new URL(strLink).openConnection());
										huc.setRequestMethod("HEAD");
										huc.connect();
										respCode = huc.getResponseCode();

										if (respCode <= 400) {
											ReportFactory.reporterOutput(strWCMTCID, testCaseDescription, inputValue,
													"Link : " + title + " should be availabe",
													linkValues.getText() + ", HTTP Respose Code : " + respCode, "Pass");
										} else {
											
											ReportFactory.reporterOutput(strWCMTCID, testCaseDescription, inputValue,
													title, "It's a Broken link and response code is : " + respCode,
													"Fail");
										}
										break;
									}

									else {

										LogFactory.info("This is not a valid AT-Link.");
										ReportFactory.reporterOutput(strWCMTCID, testCaseDescription, inputValue, title,
												linkValues.getText(), "Fail");

									}

									break;
									case "AT-Document":
									if (linkExtentionType.contains("exe") || linkExtentionType.contains("doc")
											|| linkExtentionType.contains("pdf") || linkExtentionType.contains("bmp")
											|| linkExtentionType.contains("jpg") || linkExtentionType.contains("jpeg")
											|| linkExtentionType.contains("html")
											|| linkExtentionType.contains("png")) {
										
										LogFactory.info("Content Type for the link " + title + " is AT-Document and file extention is " + linkExtentionType);

										ReportFactory.reporterOutput(strWCMTCID, testCaseDescription, inputValue,
												"Content Type should be AT-Document : " + linkExtentionType,
												title + "\n" + strLink, "Pass");

									}

									else {
										LogFactory.info(
												"This AT-Document but extension is not found in the extension maping list ");

										ReportFactory.reporterOutput(strWCMTCID, testCaseDescription, inputValue,
												"Content Type is AT-Document but:  " + linkExtentionType
														+ " not found in the extension list ",
												strLink, "Fail");

									}
									break;

								case "AT-Rich Text":

									if (linkExtentionType.contains("rtf")) {

										LogFactory.info("Content Type for the link " + title + " is AT-Rich Text and file extention is " + linkExtentionType);
										ReportFactory.reporterOutput(strWCMTCID, testCaseDescription, inputValue,
												"Content Type should be AT-RichText : " + linkExtentionType, strLink,
												"Fail");
									} else {
										LogFactory.info("Content Type for the link " + title + " is AT-Rich Text and but the file extention ( " + linkExtentionType + " ) is invalid ");
										ReportFactory.reporterOutput(strWCMTCID, testCaseDescription, inputValue,
												"Content Type should be AT-RichText : " + linkExtentionType,
												"Content Type is AT-RichText but  .rtf is not found in the href ",
												"Fail");

									}

									break;

								default:
									ReportFactory.reporterOutput(strWCMTCID, testCaseDescription, inputValue, title,
											"Invalid Content Type " + contentType, "Pass");
								}

							}
						}

					}

					// for loop for links

					if (booCountry == true && booProduct == true && booTitleFlag == false) {

						// LogFactory.info(title + " : Link is not available in the department : " +
						// departmentName + " and section : " + secondLevel);

						ReportFactory.reporterOutput(strWCMTCID, testCaseDescription, inputValue,
								title + " : Link should be available in the department : " + departmentName
										+ " and section " + secondLevel,
								"User Country & Product Types are matching with WCM Country & Product Types but Link : "
										+ title + " is not available in section : " + secondLevel,
								"Fail");

					} else if (booCountry == true && booProduct == false && booTitleFlag == true) {

						// LogFactory.info(title + " : Link is available in the department : " +
						// departmentName + " and section : " + secondLevel);

						ReportFactory.reporterOutput(strWCMTCID, testCaseDescription, inputValue,
								title + " : Link should NOT be available in the department : " + departmentName
										+ " and section " + secondLevel,
								"Uesr Product types are not matching with WCM Product types but Link : " + title
										+ " is available in section : " + secondLevel,
								"Fail");

					} else if (booCountry == false && booProduct == true && booTitleFlag == true) {

						ReportFactory.reporterOutput(strWCMTCID, testCaseDescription, inputValue,
								title + " : Link should NOT be available in the department : " + departmentName
										+ " and section " + secondLevel,
								"Uesr Country is not matching with WCM country but Link : " + title
										+ " is available in section : " + secondLevel,
								"Fail");

					} else if (booCountry == false && booProduct == false && booTitleFlag == false) {

						ReportFactory.reporterOutput(strWCMTCID, testCaseDescription, inputValue,
								title + " : Link should NOT be available in the department : " + departmentName
										+ " and section " + secondLevel,
								"Uesr Country & Product types are not matching with WCM Country & Product types so Link : "
										+ title + " is NOT available in section : " + secondLevel,
								"Pass");

					} else if (booCountry == true && booProduct == false && booTitleFlag == false) {

						ReportFactory.reporterOutput(strWCMTCID, testCaseDescription, inputValue,
								title + " : Link should NOT be available in the department : " + departmentName
										+ " and section " + secondLevel,
								"Product types are not matching with WCM Product types so Link : " + title
										+ " is not available in section : " + secondLevel,
								"Pass");

					} else if (booCountry == false && booProduct == false && booTitleFlag == true) {

						ReportFactory.reporterOutput(strWCMTCID, testCaseDescription, inputValue,
								title + " : Link should NOT be available in the department : " + departmentName
										+ " and section " + secondLevel,
								"Uesr Country & Product types are NOT matching with WCM Country & Product types but Link : "
										+ title + " is available in section : " + secondLevel,
								"Fail");

					} else if (booCountry == false && booProduct == true && booTitleFlag == false) {

						ReportFactory.reporterOutput(strWCMTCID, testCaseDescription, inputValue,
								title + " : Link should NOT be available in the department : " + departmentName
										+ " and section " + secondLevel,
								"Uesr Country is not matching with WCM country so Link : " + title
										+ " is NOT available in section : " + secondLevel,
								"Pass");

					}
				}
				// Section Block

			}
			if (sectionFlag != true) {
				LogFactory.info(secondLevel + " : section is not available in the department : " + departmentName);
				ReportFactory.reporterOutput(strWCMTCID, testCaseDescription, inputValue,
						secondLevel + " : section should be available in the department : " + departmentName,
						"Section " + secondLevel + " is not available in " + departmentName, "Fail");
			}

		} catch (

		Exception e) {

			ReportFactory.reporterOutput("Portlet Links", testCaseDescription, inputValue, "NA", e.getMessage(),
					"Fail");

		}
	}

	public static ArrayList<String> getLeftWindowLinks() {
		List<WebElement> webElements = wbelLeftWindow.findElements(By.tagName("a"));
		int webElementsSize = webElements.size();
		HashMap<Integer, WebElement> hm = new HashMap<>();
		for (int i = 1; i < webElementsSize; i++) {
			hm.put(i + 1, webElements.get(i));
		}
		List<String> strList = new ArrayList<>();
		for (Entry<Integer, WebElement> entry : hm.entrySet()) {
			strList.add(entry.getValue().getText());
		}
		return (ArrayList<String>) strList;

	}

	/*
	 * public static void verifyleftnavigationwindow(String rawData) throws
	 * Throwable { ArrayList<String> windowLinks =
	 * Homepage_POF.getLeftWindowLinks();
	 * 
	 * LogFactory.info(" List of navigation links : " + windowLinks);
	 * 
	 * List<String> ExpectedData = GenericFactory.splitString(rawData, ","); for
	 * (int i = 0; i < ExpectedData.size(); i++) { String Actualvalue =
	 * windowLinks.get(i).toString().trim(); if
	 * (Actualvalue.equals(ExpectedData.get(i))) {
	 * 
	 * LogFactory.info("Actual Value :" + Actualvalue + " || Expected value :" +
	 * ExpectedData.get(i)); //
	 * ReportFactory.ReporterOutput("STEP1","Verify_navigationlinks_homepage", //
	 * "Navigation links" , "Navigation Links should be // "+windowLinks,"Navigation
	 * links should be displayed","Pass");
	 * 
	 * } else { LogFactory.info("Actual Value :" + Actualvalue +
	 * " || Expected value :" + ExpectedData.get(i));
	 * ReportFactory.reporterOutput("STEP1", "Verify_navigationlinks_homepage",
	 * "Navigation links", "Navigation Links should be " + windowLinks,
	 * "Navigation links are not displayed", "Fail"); }
	 * 
	 * } ReportFactory.reporterOutput("STEP1", "Verify_navigationlinks_homepage", ""
	 * + windowLinks, "Navigation Links should be " + ExpectedData,
	 * "Navigation links are  displayed", "Pass");
	 * 
	 * }
	 */

}
