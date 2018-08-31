/* 
 * Project    : DealerPath
 * Script     : Alerts_POF
 * Author     : Neeraja Mantri
 * Date       : May.15.2018
 * Last Modified On:
 * Modified By :
 */

package com.deere.PageFactory;

import java.net.HttpURLConnection;
import java.net.URL;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.NoSuchElementException;

import javax.xml.xpath.XPath;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.logging.Log;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.How;
import org.testng.Assert;

import com.deere.Helpers.BaseClass;
import com.deere.Helpers.DateFactory;
import com.deere.Helpers.GenericFactory;
import com.deere.Helpers.LogFactory;
import com.deere.Helpers.ReportFactory;
import com.deere.Helpers.ValidationFactory;
import com.deere.Helpers.WaitFactory;

public class Alerts_POF {

	final WebDriver alrtDriver;

	public Alerts_POF(WebDriver driver) {
		this.alrtDriver = driver;

	}

	/*
	 * @FindBy (how=How.XPATH, using
	 * ="//h3[@class='section-title'and contains(text(),' DealerPath Alerts')]")
	 * static WebElement alertheader;
	 */

	@FindBy(how = How.XPATH, using = "//div[@class='section warning']")
	public static WebElement wbelAlertPortlet;

	@FindBy(how = How.XPATH, using = "//div[@class='section warning']//div[@class='section-header']")
	static WebElement wbelPortletHeader;
	

	@FindBy(how = How.XPATH, using = "//span[@class='icon warning']")
	static WebElement wbelWarningImg;

	@FindBy(how = How.XPATH, using = ".//*[@id='layoutContainers']/div/div[2]/div[3]/div[2]/section/div/div[2]")
	public static WebElement wbelAlertFramePath;

	@FindBy(how = How.XPATH, using = "//div[@class='section warning']//span[@class='list-item-title']")
	public static WebElement wbeAlertsTitle;

	@FindBy(how = How.XPATH, using = ".//*[@id='layoutContainers']/div/div[2]/div[3]/div[2]/section/div/div[2]/div")
	static WebElement alertIndex;

	@FindBy(how = How.XPATH, using = ".//*[@id='layoutContainers']/div/div[2]/div[3]/div[3]/section/div/div[2]/div")
	static WebElement announcementIndex;

	@FindBy(how = How.XPATH, using = ".//*[@id='layoutContainers']/div/div[2]/div[3]/div[4]/section/div/div[2]/div")
	static WebElement favoritesIndex;

	@FindBy(how = How.XPATH, using = ".//*[@id='layoutContainers']/div/div[2]/div[3]/div[1]")
	static WebElement searchIndex;
	
	@FindBy(how = How.XPATH, using = ".//*[@id='layoutContainers']/div/div[2]/div[3]/div[4]/section/div/div[4]")
	static WebElement linkPortletIndex;
	
	@FindBy(how = How.XPATH, using = "//div[@class='section-header']/h3")
	static List<WebElement> sectionHeader;

	@FindBy(how = How.XPATH, using = "//div[@class='section warning']/div[2]/div//div/div[1]/span")
	static List<WebElement> AlertHeaderTitle;

	@FindBy(how = How.XPATH, using = ".//div[@class='section warning']//div[@class='list-item-body']/span")
	public static List<WebElement> alertsTableTitle;

	@FindBy(how = How.XPATH, using = ".//*[@id='layoutContainers']/div/div[2]/div[3]/div[2]/section/div/div[2]/div[2]/div/div[6]/div[2]/div")
	public static WebElement readMoreLink;
	
	@FindBy(how = How.XPATH, using = "//a[@class='active leftNavIndexPadding ']")
	public static List<WebElement> SecondLevelDepartment;
	
	@FindBy(how = How.XPATH, using = ".//*[@id='js-segments-popover']/div[2]/div[3]/button")
	public static WebElement ApplyFilterButton;
	
	@FindBy(how = How.XPATH, using = "//div[@class='wide-list hide-overflow is-expanded']")
	public static WebElement Announcementportlet;

	@FindBy(how = How.XPATH, using = ".//*[@id='layoutContainers']/div/div[2]/div[3]/div[4]/section/div/div[4]")
	public static WebElement Linksportlet;
	
	@FindBy(how = How.XPATH, using = "//*[@id='productSegmentsError']")
	public static WebElement wblProdSegmentError;


	/**
	 * @throws Throwable
	 *//*
	public static void isAlertPortletPresent() throws Throwable {

		String strFlag = "FAIL";
		String strResult = "Alerts are not displayed";
		try {
			if (ValidationFactory.isElementPresent(wbelAlertPortlet)) {
				strFlag = "PASS";
				strResult = "Alerts are displayed";
			}
			ReportFactory.reporterOutput("TC05_Homepage", "Verify presence of Alert portlet on Home Page", "NA","Alert should be present on homepage", strResult, strFlag);
			
			if (strFlag.equalsIgnoreCase("FAIL") )
			{ 	Assert.assertFalse(true);}

		} catch (Throwable e) {

			ReportFactory.reporterOutput("TC05_Homepage", "Verify presence of Alert portlet on Home Page", "NA","Alert should be present on homepage", e.getMessage().toString(), strFlag);
		}

	}*/
	
	/**
	 * @throws Throwable
	 */
	public static String isAlertPortletPresent() throws Throwable {

		String strFlagForAlertPortletPresence = "Fail";
		String strResult = "Alerts are not displayed";
		try {
			if (ValidationFactory.isElementPresent(wbelAlertPortlet)) {
				strFlagForAlertPortletPresence = "Pass";
				strResult = "Alerts are displayed";

			}

			else {

				LogFactory.info("Alert portlet not present");

			}

			return strFlagForAlertPortletPresence;
		} catch (Throwable e) {

			LogFactory.info("Exception :-->" + e.getMessage().toString());
			return strFlagForAlertPortletPresence;
		}

	}

	
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

		} catch (Throwable e) {

			ReportFactory.reporterOutput(TCID, "Verify alert warning sign icon on homepage", "NA",
					"Alert warning sign icon should be present on homepage", e.getMessage().toString(), strFlag);
		}
	}

	/**
	 * @param Title
	 * @return
	 * @throws Throwable
	 */
	public static boolean alertTitleComparison(String Title) throws Throwable {

		boolean booAlertsTitle = false;
		String flag = "Fail";
		String result = "Alerts are not present";

		// List<String> expectedTitleFromExcel = GenericFactory.splitString(Title, ",");

		List<String> actualTitle = getListOfAlertsWithTitle();

		if (isAlertPortletPresent().equalsIgnoreCase("Pass"))
		{
			/* Get the title of alert and verify with the testdata title */

			if (actualTitle.contains(Title)) {

				LogFactory.info("Title is present in alerts title list");
				flag = "Pass";
				result = "Alerts are present for the user";
				booAlertsTitle = true;

			} else {

				LogFactory.info("Title is not present in alerts title list");

			}

		}
		//ReportFactory.reporterOutput("Alerts", "Verify Alerts presence on Homepage", Title,"Alerts should be present on the site", result + " " + Title, flag);
		return booAlertsTitle;

	}

	/**
	 * @return
	 */
	public static ArrayList<String> getListOfAlertsWithTitle() {

		List<WebElement> numberOfAlerts = BaseClass.wbDriver.findElements(By.xpath("//div[@class='section warning']"));
		//System.out.println("No of alerts are :" + numberOfAlerts.size());

		List<WebElement> liAlertsTitle = BaseClass.wbDriver
				.findElements(By.xpath("//div[@class='section warning']//span[@class='list-item-title']"));
		//System.out.println("Number of alerts with titles :" + liAlertsTitle.size());

		ArrayList<String> alertsWithTitles = new ArrayList<String>();

		for (int i = 0; i < liAlertsTitle.size(); i++) {

			String strAlertTitleFromSite = liAlertsTitle.get(i).getText().toString();

			String strAlertTitle = StringUtils.substringAfter(strAlertTitleFromSite, " :").trim();
			// String strAlertTitle=strAlertTitleFromSite.split(":").toString();
			//System.out.println("Alert with title :" + strAlertTitle);

			alertsWithTitles.add(strAlertTitle);

		}

		//System.out.println("List with alert titles :" + alertsWithTitles);

		return alertsWithTitles;

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
		String strResult = " Alert header Text is not displayed in the user preffered language";
		try {
			if (ValidationFactory.isElementPresent(wbelPortletHeader)) {

				LogFactory.info("Alerts portlet header is present ");
				
				strAlrtHeaderTxtOnHomePage = wbelPortletHeader.getText();
				if (strAlrtHeaderTxtOnHomePage.equals(strAlertHeaderTxtPrefrdLang)) {
					strFlag = "PASS";
					LogFactory.info("Alert is displayed in the user  preffered language with " + strAlrtHeaderTxtOnHomePage + " "
							+ "" + strAlertHeaderTxtPrefrdLang);
					strResult = " Alert header text is displayed in the user preferred language";
				}
				ReportFactory.reporterOutput(strTCID, "Verify alert header text in the user preffered language",
						strAlrtHeaderTxtOnHomePage, "Alert header text should be displayed in the user preffered language",
						strResult, strFlag );

			}

		} catch (Throwable e) {

			ReportFactory.reporterOutput(strTCID, "Verify alert header text in the user preffered language",
					strAlrtHeaderTxtOnHomePage, "Alert header text should be displayed in preffered language",
					e.getMessage().toString(), strFlag );
		}

	}
	
	/**
	 * @author Neeraja.mantri
	 * @createdAt 03-07-2018
	 * @param UserCountry
	 * @param wcmCountry
	 * @param userProductList
	 * @param wcmProductList
	 * @param strDepartmentName
	 * @param strSecondlevelDepartment
	 * @param strIndexPage
	 * @param strTCID
	 * @throws Throwable
	 */
	public void checkAlertsPresenceOnFirstIndexPage(String UserCountry, String wcmCountry, List<String> userProductList,
			List<String> wcmProductList, String strDepartmentName,String strSecondlevelDepartment,String strIndexPage,String strTCID) throws Throwable 
	{
		String strFlag="";
		String strResult="";
		boolean booIndexFlag=false;
		boolean booProductMatchFlag=false;
		boolean flag= false;
		try {
		for(int i=0; i< userProductList.size();i++) 
		{
			List<String> userProducts = GenericFactory.strUserProductToMatchWithWCMList(userProductList.get(i));
			
			booProductMatchFlag=GenericFactory.verifyIfProductMatch(wcmProductList, userProducts);
			if(booProductMatchFlag==true)
			{
				break;
			}
		}
		if(GenericFactory.userAndWCMCountryComparison(UserCountry, wcmCountry)==true &&  booProductMatchFlag==true )
		{
			
		LogFactory.info("User country and product types matches with wcm country and product types");
			
			WebElement department=GenericFactory.getDeptname(strDepartmentName);
	   	department.click();
	    WaitFactory.waitForPageLoaded();
	   
	    LogFactory.info("Clicked on enabled department");
		
	   
	    WebElement  secondLevelDepartment=GenericFactory.getActive2ndLevelDepartment(strSecondlevelDepartment);
		   secondLevelDepartment.click();
		   
		   WaitFactory.waitForPageLoaded();
		   
		   LogFactory.info("Second level department :"+ secondLevelDepartment.getText()+" is clicked");
		   if(GenericFactory.clickOnIndexPageLinkOnSecondLevelDepartment(strSecondlevelDepartment, strIndexPage))
		   {
			   WaitFactory.waitForPageLoaded();
			   if (isAlertPortletPresent().equalsIgnoreCase("Fail"))
		
			   {
				   LogFactory.info("Alerts portlet is not present on 3rd level index page" );
				  // booThirdLevelIndex=true;
				   strFlag="Pass";
				   booIndexFlag=true;
				   strResult="Alerts portlet is not present on third level index page : "+strIndexPage +" under department"+strDepartmentName+" and second level department "+strSecondlevelDepartment;
			   }
			   else {
				   LogFactory.info("Alerts portlet is present on 3rd level index page" );
				   strFlag="Fail";
				   strResult="Alerts portlet is present on third level index page : "+strIndexPage +" under department "+strDepartmentName+" and second level department "+strSecondlevelDepartment;
			      }
	   }
		   else {
			   
			   LogFactory.info("Required index page is not present");
			   strFlag="Fail";
			   strResult="Required index page "+strIndexPage+"is not present under department "+strDepartmentName+ "and second level department "+strSecondlevelDepartment; 
		   }   
		   
		  // ReportFactory.reporterOutput("Verify alerts portlet not available on index pages",strTCID, "Department name :" +strDepartmentName +" and secondlevel department is: "+strSecondlevelDepartment,"Alerts portlet should not be present on index pages",strResult, strFlag);  
		Homepage_POF.homepagepath.click();   
	}		   
		else {
			LogFactory.info("Country and product types are not matching");
			strResult="Country and product types of the dealer are not matching";
			strFlag="Fail";
			
		}
		   
		
		
		 ReportFactory.reporterOutput(strTCID,"Verify alerts portlet not available on index pages", "Department name :" +strDepartmentName +" and secondlevel department is: "+strSecondlevelDepartment,"Alerts portlet should not be present on index pages",strResult, strFlag);  
		
		
		
		}
		catch (Exception e) {
				LogFactory.info(e.getMessage());
			
		}
		
		
		
	}	
	
	/**
	 * @author Neeraja.mantri
	 * @createdAt 03-07-2018
	 * @param userCountry
	 * @param wcmCountry
	 * @param UserProductName
	 * @param WcmProducts
	 * @param titleName
	 * @param strTCID
	 * @throws Throwable
	 */
	public static void checkAlertContentFilterOnProductTypes(String userCountry,String wcmCountry,String UserProductName, String WcmProducts,String titleName,String strTCID) throws Throwable
	{
		String strAlertFlag="";
		String strResult="";
	
		if (alertTitleComparison(titleName)==true) 
		{
			GenericFactory.checkAllProducts();
			
			LogFactory.info("Alerts portlet is present");
			List<String> listUserProducts = GenericFactory.splitString(UserProductName, ",");
			List<String> listWcmProducts = GenericFactory.splitString(WcmProducts, ",");
			
			ProductSegment_POF.wbelProductSegmentIcon.click();
			 
			for (int i = 0; i < listUserProducts.size(); i++) 
			{
				 List<String> listUserProductstype = GenericFactory.strUserProductToMatchWithWCMList(listUserProducts.get(i));
				 if(GenericFactory.verifyIfProductMatch(listWcmProducts, listUserProductstype))
				 {
					 LogFactory.info("User products and wcm products are matching");
					 GenericFactory.productUncheck((GenericFactory.getParentProductSegmnent(listUserProducts.get(i))));
				 }	
			}	
			ApplyFilterButton.click();
			if (!ValidationFactory.isElementPresent(wblProdSegmentError)) {
						
			if(!alertTitleComparison(titleName)) {
					 	strAlertFlag="Pass"; 
					 	strResult="Alert title :" +titleName+ "is not present after the corresponding product types are filtered " +WcmProducts ;
				 }
				 else {
					 
					 	strAlertFlag="Fail";
					 	strResult="Alert title is still present even after applying product type filters";
					 }
			
			}else  
			{	
				strResult="Unable to perform Product Type filters verification due to " + wblProdSegmentError.getText();
				strAlertFlag="Pass";
				ProductSegment_POF.wbelProductSegmentIcon.click();
			}
		
			ReportFactory.reporterOutput(strTCID,"Verify Alert content on changing of preferred product types", titleName,
					"Alerts contents should filter based on product types filters",strResult, strAlertFlag);

			
	}
	}
	
	/*************************
	 * Check alert header and text present for homepage and department
	 *******************************************/
	/**
	 * @author Neeraja.mantri
	 * @createdAt 22-06-2018
	 * @param actualheadertxt
	 * @param TCID
	 * @throws Throwable
	 */
	public static boolean getAlertPortletHeaderAndText(String strAlertHeaderTextFromExcel,String strDepartmentNameFromExcel, String strTCID) throws Throwable {

		String strFlag = "Fail";
		String strResult = "";
		String strResultPresentOnDepartment="";
		String strResultNotPresentOnDepartment="";
		
		boolean booFlagAvailableOnDepartment=false;
		boolean booFlagNotAvailableOnDepartment=false;
		
		List<String> liDepartmentNameFromExcel = GenericFactory.splitString(strDepartmentNameFromExcel, ",");
		boolean booFlagAlertHeaderAndText = false;
		
		List<String> liMatchingDepartments=new ArrayList<String>();
		String strMatchingDepartments="" ;
		
		List<String> liNonMatchingDepartments=new ArrayList<String>();
		String strNonMatchingDepartments="" ;
		
		try {
			if(!strDepartmentNameFromExcel.equalsIgnoreCase("NA")) {
			if (liDepartmentNameFromExcel.size() > 0) {
				for (int j = 0; j < liDepartmentNameFromExcel.size(); j++) {
					WebElement department = GenericFactory.getDeptname(liDepartmentNameFromExcel.get(j));
					if (department != null) {

						String strNewDepartmentValue=liDepartmentNameFromExcel.get(j);
						
						department.click();
					if (ValidationFactory.isElementPresent(wbelPortletHeader)) 
						{

							String strAlertHeaderTextFromSite = wbelPortletHeader.getText();
							if (strAlertHeaderTextFromSite.equalsIgnoreCase(strAlertHeaderTextFromExcel))
							{
								LogFactory.info("ALert header text is present on department page" + " : "+ liDepartmentNameFromExcel.get(j));
								
							
								liMatchingDepartments.add(liDepartmentNameFromExcel.get(j));

								//strMatchingDepartments = String.join(",", liMatchingDepartments);
								
								//booFlagAvailableOnDepartment=true;
								booFlagAlertHeaderAndText = true;
							}

							else {

								LogFactory.info("ALert header text is not present on department page" + " : "+ liDepartmentNameFromExcel.get(j));
								
								booFlagNotAvailableOnDepartment=true;
								liNonMatchingDepartments.add(liDepartmentNameFromExcel.get(j));
								
							//	strNonMatchingDepartments = String.join(",", liNonMatchingDepartments);
								
								strResultNotPresentOnDepartment=strResultNotPresentOnDepartment.concat(strNewDepartmentValue);
						}
					} else 
						{

						LogFactory.info("ALert portlet is not present on department page" + " : "+ liDepartmentNameFromExcel.get(j));
							
						booFlagNotAvailableOnDepartment=true;
						
						strResultNotPresentOnDepartment=strResultNotPresentOnDepartment.concat(strNewDepartmentValue);
							
							liNonMatchingDepartments.add(liDepartmentNameFromExcel.get(j));
							// strNonMatchingDepartments = String.join(",", liNonMatchingDepartments);
							
						}
						Homepage_POF.homepagepath.click();
				}
						
				}
				
				//if(!booFlagAvailableOnDepartment || booFlagNotAvailableOnDepartment)
					if(liMatchingDepartments.size()>0 )
				{
					strFlag = "Pass";
					LogFactory.info("Alert header text is present on active departments : "+liMatchingDepartments);
					strResult="Alert header text is present on active departments :"+liMatchingDepartments;
					
					ReportFactory.reporterOutput(strTCID,"Verify alerts header text is present on department page", strDepartmentNameFromExcel,"Alert header text should be displayed on active department pages", strResult, strFlag);
			
				}
					//else if(booFlagAvailableOnDepartment || !booFlagNotAvailableOnDepartment)
					else if(liMatchingDepartments.size()==0 || liNonMatchingDepartments.size()>0)
				{
					strFlag = "Fail";
					
					LogFactory.info("Alert header text is present on active departments : "+liMatchingDepartments + "and not present on departments :"+liNonMatchingDepartments);
					
					strResult="Alert header text is present on active departments :"+liMatchingDepartments + "and not present on departments :"+liNonMatchingDepartments;
					
					ReportFactory.reporterOutput(strTCID,"Verify alerts header text is present on department page", strDepartmentNameFromExcel,"Alert header text should be displayed on  active department pages", strResult, strFlag);
			
					
				}
				
			}
			}
			else {

				LogFactory.info("Verifying alert header text is displayed on homepage");
			
				if (ValidationFactory.isElementPresent(wbelPortletHeader)) {

					String strAlertHeaderTextFromSite = wbelPortletHeader.getText();
					if (strAlertHeaderTextFromSite.equalsIgnoreCase(strAlertHeaderTextFromExcel)) {
						LogFactory.info("ALert header text is present on homepage");
						strFlag = "Pass";
						strResult = "Alert header text is displayed on homepage";
						ReportFactory.reporterOutput(strTCID,"Verify alerts header text is present on homepage", 
								strAlertHeaderTextFromExcel, "Alert header text should be displayed on homepage",
								strResult, strFlag);
						booFlagAlertHeaderAndText = true;
					} else {

						LogFactory.info("ALert header text is not present on homepage");
						strResult = "Alert header text is not displayed on homepage";
						ReportFactory.reporterOutput( strTCID,"Verify alerts header text is present on homepage",
								strAlertHeaderTextFromExcel, "Alert header text should be displayed on homepage",
								strResult, strFlag);
					}

				}

			}
			
		} catch (Throwable e) {

			LogFactory.info("Verify alerts header text is present on homepage" + e.getMessage().toString());

		}

		return booFlagAlertHeaderAndText;
		
		}
	
	
	/*************************
	 * Check for warning sign present in alert header
	 *******************************************/

	/**
	 * @author Neeraja.mantri
	 * @createdAt 22-05-2018
	 * @param strDepartmentNameFromExcel
	 * @param strTCID
	 * @return
	 * @throws Throwable
	 */

	public static boolean checkAlertIconDisplayedInHeader(String strDepartmentNameFromExcel, String strTCID)
			throws Throwable

	{
		String strFlag = "Fail";
		String strResult = "";
		boolean booFlagAlertImage = false;
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
					if (department != null)
					{
						String strNewDepartmentValue=liDepartmentNameFromExcel.get(j);
						department.click();
						if (ValidationFactory.isElementPresent(wbelWarningImg)) 
						{
							LogFactory.info("ALert icon is present on active department page" + " : "+ liDepartmentNameFromExcel.get(j));
							liMatchingDepartments.add(liDepartmentNameFromExcel.get(j));
							strResultPresentOnDepartment=strResultPresentOnDepartment.concat(strNewDepartmentValue);
							
							booFlagAlertImage = true;
						}
						else {
								LogFactory.info("ALert icon is not present on department page" + " : "+ liDepartmentNameFromExcel.get(j));
								liNonMatchingDepartments.add(liDepartmentNameFromExcel.get(j));
						}
						Homepage_POF.homepagepath.click();
					}
				}
				if(liMatchingDepartments.size()>0 )
				{
					strFlag = "Pass";
					LogFactory.info("Alert icon is present on active departments : "+liMatchingDepartments);
					strResult="Alert icon is present on active departments  :"+liMatchingDepartments;
					
					ReportFactory.reporterOutput(strTCID,"Verify alerts icon is present on department page", strDepartmentNameFromExcel,"Alerts icon should be present on active department page", strResult, strFlag);
			
				}
					
					else if(liMatchingDepartments.size()==0 || liNonMatchingDepartments.size()>0)
				{
					strFlag = "Fail";
					
					LogFactory.info("Alert header text is present on active departments : "+liMatchingDepartments + "and not present on departments :"+liNonMatchingDepartments);
					
					strResult="Alert header text is not present on departments :"+liNonMatchingDepartments;
					
					ReportFactory.reporterOutput(strTCID,"Verify alerts icon is present on department page", strDepartmentNameFromExcel,"Alerts icon should be present on active department page", strResult, strFlag);
			
				}
			} 
			}
			else {
				LogFactory.info("Verifying alert icon is displayed on homepage");

				strFlag = "Fail";
				if (ValidationFactory.isElementPresent(wbelWarningImg)) {
					LogFactory.info("Alert icon is present on homepage");
					strFlag = "Pass";
					strResult = "Alert icon is displayed on homepage";
					ReportFactory.reporterOutput(strTCID,"Verify icon is present on homepage", 
							"NA", "Alerts icon should be present on homepage", strResult, strFlag);
					booFlagAlertImage = true;
				} else {

					LogFactory.info("Alert icon is not present on homepage");
					strResult = "Alert icon is not displayed on homepage";
					ReportFactory.reporterOutput(strTCID,"Verify alerts icon is present on homepage", 
							"NA", "Alerts icon should be present on homepage", strResult, strFlag);
				}

			
		}
		} catch (Throwable e) {
			LogFactory.info("Verify alerts icon is present on homepage" + e.getMessage().toString());
			
		}

		return booFlagAlertImage;
	}
	
	/*********************
	 * check portlets order on homepage and department
	 ****************************************/

	/**
	 * @author Neeraja.mantri
	 * @createdAt 22-05-2018
	 * @param strDeptName
	 * @param strTCID
	 * @return
	 * @M
	 */
	/*
	public static boolean checkOrderOfPortletsOnHomepage (String strDepartmentNamesFromExcel, String strPortletHeaders, String strTCID) 
	{
		String strFlag = "Fail";
		String strResult = "";
		String strResultPresentOnDepartment="";
		String strResultNotPresentOnDepartment="";
		
		List<String> liDepartmentNameFromExcel = GenericFactory.splitString(strDepartmentNamesFromExcel, ",");
		boolean booFlagCheckPortletOrder = false;
		
		
		try {
			if(!strDepartmentNamesFromExcel.equalsIgnoreCase("NA")) {
			if (liDepartmentNameFromExcel.size() > 0) {
				
				LogFactory.info("Check portlet order for department");
			/*	for (int j = 0; j < liDepartmentNameFromExcel.size(); j++) {
					WebElement department = GenericFactory.getDeptname(liDepartmentNameFromExcel.get(j));
					if (department != null) {
						
						String strNewDepartmentValue=liDepartmentNameFromExcel.get(j);
						department.click();
						LogFactory.info(
								"Checking portlets order for department" + " : " + liDepartmentNameFromExcel.get(j));
						checkorderofportletsdisplayed(strPortletHeaders,strTCID);

						booFlagCheckPortletOrder = true;

					}
					Homepage_POF.homepagepath.click(); */
			/*	}
			}
			else {
				LogFactory.info(
						"Department is not present for checkOrderOfPortletsOnHomepage&Department method and is on homepage");
				checkorderofportletsdisplayed(strPortletHeaders,strTCID);
				booFlagCheckPortletOrder = true;

			}

		} catch (Throwable e) {

			LogFactory.info("Exception :" + e.getMessage().toString());

		}

		return booFlagCheckPortletOrder;
	}  */

	
	/**
	 * @param strTCID
	 * @throws Throwable
	 */
	public static void checkorderofportletsdisplayed(String strPortletHeaders,String strTCID) throws Throwable {
		String flag = "Fail";
		String resultPortletOrder = "";

		// List adding portlet frames
		ArrayList<WebElement> listPortletframes = new ArrayList<WebElement>();
		listPortletframes.add(searchIndex);
		listPortletframes.add(alertIndex);
		listPortletframes.add(announcementIndex);
		listPortletframes.add(favoritesIndex);

		int intListPortletFrameSize = listPortletframes.size();
		List<String> headerTxtExcelValues = GenericFactory.splitString(strPortletHeaders, ",");

		HashMap<Integer, String> headerTxtHomePage = new HashMap<Integer, String>();
		int intListPageSize = sectionHeader.size();

		for (int i = 0; i < sectionHeader.size(); i++) {
			headerTxtHomePage.put(i, sectionHeader.get(i).getText());

			if (sectionHeader.get(i).getText().contains(headerTxtExcelValues.get(1))) {

				String strAnnouncementHeaderTxt = sectionHeader.get(i).getText();
				String strAnnouncementTxt = strAnnouncementHeaderTxt.substring(0,
						strAnnouncementHeaderTxt.lastIndexOf('(') - 1);

				headerTxtHomePage.put(i, strAnnouncementTxt);
			}

			// System.out.println("List of elements :" + sectionHeader.get(i).getText());
		}
		int headerValue = headerTxtHomePage.size();

		switch (headerValue) {
		case 3:
			resultPortletOrder = "Portlets are not displayed as expected";
			for (int i = 0; i < headerTxtExcelValues.size(); i++) {
				if (headerTxtExcelValues.get(i).equals(headerTxtHomePage.get(i))) {
					flag = "Pass";
					// resultPortletOrder="Alerts,Announcements & favorites portlet order are
					// displayed as expected";
					resultPortletOrder = "Alerts portlet present above announcements & announcements portlet present above favorites portlet as expected";
					LogFactory
							.info(headerTxtExcelValues.get(i) + " " + "are equal" + " " + headerTxtExcelValues.get(i));
				}

				else {
					LogFactory.info(
							headerTxtExcelValues.get(i) + " " + "are not equal" + " " + headerTxtExcelValues.get(i));
					flag = "Fail";
					break;
				}

			}
			LogFactory.info("All the portlets are present in order");
			ReportFactory.reporterOutput(strTCID,"Verify order of portlets displayed",  "NA",
					"Portlets should be displayed in order :Alert portlet in first,announcement portlet in second and favorites portlet in last",
					resultPortletOrder, flag);
			break;
		case 2:

			resultPortletOrder = "Portlets are not displayed as expected";

			for (int i = 0; i < headerTxtHomePage.size(); i++) {

				String strHeaderTxt = headerTxtHomePage.get(i);
				if (strHeaderTxt.equals("DealerPath Alerts")) {
					flag = "Pass";
					resultPortletOrder = "Alert portlet is displayed first and above announcements & favorites portlet";
					LogFactory.info("ALerts portlet is displayed first ");
					ReportFactory.reporterOutput(strTCID,"Verify order of portlets displayed",  "NA",
							"Portlets should be displayed in order :Alert portlet in first,announcement portlet in second and favorites portlet in last",
							resultPortletOrder, flag);
					// break;
				} else if (strHeaderTxt.equals("My DealerPath Announcements")) {
					flag = "Pass";
					resultPortletOrder = "Announcement portlet is displayed as expected and above favorites portlet or below alerts portlet";
					LogFactory.info("Announcements portlet is displayed");
					ReportFactory.reporterOutput(strTCID,"Verify order of portlets displayed",  "NA",
							"Portlets should be displayed in order :Alert portlet in first,announcement portlet in second and favorites portlet in last",
							resultPortletOrder, flag);
					break;
				} else if (strHeaderTxt.equals("My DealerPath favorites")) {
					flag = "Pass";
					LogFactory.info("favorites portlet is displayed");
					resultPortletOrder = "Favorites portlet is displayed at last";
					ReportFactory.reporterOutput(strTCID,"Verify order of portlets displayed",  "NA",
							"Portlets should be displayed in order :Alert portlet in first,announcement portlet in second and favorites portlet in last",
							resultPortletOrder, flag);
					break;
				} else {
					flag = "Fail";
					ReportFactory.reporterOutput(strTCID,"Verify order of portlets displayed",  "NA",
							"Portlets should be displayed in order :Alert portlet in first,announcement portlet in second and favorites portlet in last",
							resultPortletOrder, flag);
					break;
				}

			}

			break;
		case 1:

			for (int i = 0; i < headerTxtHomePage.size(); i++) {
				String header = headerTxtHomePage.get(i);
				LogFactory.info("Only" + header + "portlet is present");
				resultPortletOrder = "Only" + header + "portlet is displayed";
				flag = "Pass";
			}
			ReportFactory.reporterOutput(strTCID,"Verify order of portlets displayed",  "NA",
					"Portlets should be displayed in order :Alert portlet in first,announcement portlet in second and favorites portlet in last",
					resultPortletOrder, flag);

			break;
		default:

			String resultOrder = "No portlets are displayed";
			LogFactory.info("No portlets are displayed");
			flag = "Fail";
			ReportFactory.reporterOutput(strTCID,"Verify order of portlets displayed",  "NA",
					"Portlets should be displayed in order :Alert portlet in first,announcement portlet in second and favorites portlet in last",
					resultOrder, flag);
		}
	}

	/**
	 * @author Neeraja.mantri
	 * @createdAt 22-06-2018
	 * @param strAlertHeaderTextFromExcel
	 * @param strDepartmentNameFromExcel
	 * @param strTCID
	 * @modifiedAt 26-06-2018
	 * @return
	 */
	/*public static boolean checkAlertsInDescendingOrderOnDateTime(String strAlertHeaderTextFromExcel,
			String strDepartmentNameFromExcel, String strTCID) {

		String strFlag = "Fail";
		

		String strResult = "";
		String strResultPresentOnDepartment="";
		String strResultNotPresentOnDepartment="";
		
		List<String> liMatchingDepartments=new ArrayList<String>();
		String strMatchingDepartments = String.join(",", liMatchingDepartments);
		
		List<String> liNonMatchingDepartments=new ArrayList<String>();
		String strNonMatchingDepartments = String.join(",", liNonMatchingDepartments);
		
		List<String> liDepartmentNameFromExcel = GenericFactory.splitString(strDepartmentNameFromExcel, ",");
		boolean booFlagAlertDisplayedDescendingOrder = false;
		try {
			if(!strDepartmentNameFromExcel.equalsIgnoreCase("NA")) {
			if (liDepartmentNameFromExcel.size() > 0) {
				for (int j = 0; j < liDepartmentNameFromExcel.size(); j++) 
				{
					WebElement department = GenericFactory.getDeptname(liDepartmentNameFromExcel.get(j));
					if (department != null) 
					{
						String strNewDepartmentValue=liDepartmentNameFromExcel.get(j);

						department.click();
					//***********
						 * Get list of alerts and verify if sorted according to date
						 ***************//*
						if (ValidationFactory.isElementPresent(wbelPortletHeader)) 
						{
							List<WebElement> numberOfAlerts = AlertHeaderTitle;
							LogFactory.info("Number of alerts are :" + numberOfAlerts.size());

							List<String> AlertsHeaderDates = GenericFactory.getListOfDatesByHeaderName(strAlertHeaderTextFromExcel);
							// System.out.println("List of dates:"+AlertsHeaderDates);

							if (GenericFactory.verifyDateSortedOrder(AlertsHeaderDates) == true) 
							{
								LogFactory.info("Alerts title dates are displayed in descending order in department page"+ " : " + liDepartmentNameFromExcel.get(j));
								
								liMatchingDepartments.add(liDepartmentNameFromExcel.get(j));
								strResultPresentOnDepartment=strResultPresentOnDepartment.concat(strNewDepartmentValue);
								
							
							} else {

								LogFactory.info("Alerts title dates are not in descending order in department page" + " : "+ liDepartmentNameFromExcel.get(j));
								
								liNonMatchingDepartments.add(liDepartmentNameFromExcel.get(j));
								
								strResultNotPresentOnDepartment=strResultNotPresentOnDepartment.concat(strNewDepartmentValue);
								
								
							}

							Homepage_POF.homepagepath.click();

						} else {

							LogFactory.info("Alert portlet is not present on page:");
							
							strResultNotPresentOnDepartment=strResultNotPresentOnDepartment.concat(strNewDepartmentValue);
							
							liNonMatchingDepartments.add(liDepartmentNameFromExcel.get(j));

						
						}
						Homepage_POF.homepagepath.click();
					}
				}
				
				if(liMatchingDepartments.size()>0 )
				{
					strFlag = "Pass";
					LogFactory.info("Alerts title dates are displayed in descending order on active departments : "+liMatchingDepartments);
					strResult="Alerts title dates are displayed in descending order on active departments :"+liMatchingDepartments;
					
					ReportFactory.reporterOutput(strTCID,"Verify alerts displayed in descending order with newest one's on top of department", strDepartmentNameFromExcel,"Alerts should be displayed in descending order with newest one's on top on active department pages", strResult, strFlag);
			
				}
				
					else if(liMatchingDepartments.size()==0 || liNonMatchingDepartments.size()>0)
				{
					strFlag = "Fail";
					
					LogFactory.info("Alerts title dates are displayed in descending order on active departments : "+liMatchingDepartments + "and not present on departments :"+liNonMatchingDepartments);
					
					strResult="Alerts title dates are not displayed in descending order on  departments  :"+liNonMatchingDepartments;
					
					ReportFactory.reporterOutput(strTCID,"Verify alerts displayed in descending order with newest one's on top of department", strDepartmentNameFromExcel,"Alerts should be displayed in descending order with newest one's on top on active department pages", strResult, strFlag);
			
				}
				}
			} else {

				LogFactory.info("Verifying alerts displaying in descending order on date and time in homepage");


				List<WebElement> numberOfAlerts = BaseClass.wbDriver.findElements(By.xpath("//div[@class='section warning']"));
				LogFactory.info("No of alerts are :" + numberOfAlerts.size());

				List<String> AlertsHeaderDates = GenericFactory.getListOfDatesByHeaderName(strAlertHeaderTextFromExcel);
				System.out.println("List of dates:" + AlertsHeaderDates);

				if (GenericFactory.verifyDateSortedOrder(AlertsHeaderDates) == true) {
					LogFactory.info(
							"Alerts title dates are in descending order with newest one's on top by date on homepage");
					strFlag = "Pass";
					strResult = "Alerts title dates are in descending order with newest one's on top by date on homepage";
					ReportFactory.reporterOutput(
							strTCID,"Verify alerts displayed in descending order with newest one's on top of homepage", 
							"NA", "Alerts should be displayed in descending order with newest one's on top of homepage",
							strResult, strFlag);
					booFlagAlertDisplayedDescendingOrder = true;
				} 
				else {

					LogFactory.info("Alerts title dates are not in sorted order on homepage");
					strResult = "Alerts title dates are not in sorted order on homepage";
					ReportFactory.reporterOutput(
							strTCID,"Verify alerts displayed in descending order with newest one's on top of homepage", 
							"NA", "Alerts should be displayed in descending order with newest one's on top of homepage",
							strResult, strFlag);
				}

			}
		} catch (Throwable e) {

			LogFactory.info("Verify alerts displayed in descending order with newest one's on top of homepage" + e.getMessage().toString());

		}

		return booFlagAlertDisplayedDescendingOrder;
	}
	*/
	
	
	
	public static void verifyAlertsShowingInDescendingOrderOnDateTime( String strTCID,String strDepartmentNameFromExcel) throws Throwable {
		
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
							if (ValidationFactory.isElementPresent(wbelAlertFramePath)) {
								List<WebElement> numberOfAnnouncement = AlertHeaderTitle;

								LogFactory.info("Number of Announcement are :" + numberOfAnnouncement.size());
								
								List<String> AnnouncementHeaderDates = GenericFactory.getListOfDatesByFrame(wbelAlertFramePath,"Alert");
								
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
						strResult="Alert title dates are displayed in descending order on active departments :"+liMatchingDepartments;
					}
					else if(liMatchingDepartments.size()==0 || liNonMatchingDepartments.size()>0){
						strFlag = "Fail";
						strResult="Alert title dates are NOT displayed in descending order on departments  :"+liNonMatchingDepartments;
					}
					ReportFactory.reporterOutput(strTCID,"Verify Alert displayed in descending order with newest one's on top on department ", strDepartmentNameFromExcel,"Alerts should be displayed in descending order with newest one's on top on active department pages", strResult, strFlag);
				}
			} 
			else {

					GenericFactory.navigateToHomePage();
					LogFactory.info("Verifying Alerts are  displaying in descending order on date and time in homepage");
					/*********** * Get list of Alerts and verify if sorted according to date ***************/
					List<WebElement> numberOfAlert = BaseClass.wbDriver.findElements(By.xpath("//div[@class='section warning']"));
					List<String> AlertHeaderDates = GenericFactory.getListOfDatesByFrame(wbelAlertFramePath,"Alert" );
		
					if (GenericFactory.verifyDateSortedOrder(AlertHeaderDates) == true) {
						
						strFlag = "Pass";
						strResult = "Alert title dates are in descending order with newest one's on top by date on homepage";
					} 
					else {
							LogFactory.info("Alert title dates are not in sorted order on homepage");
							strResult = "Alert title dates are not in sorted order on homepage";
					}
					ReportFactory.reporterOutput(
							strTCID,"Verify Alerts are displayed in descending order with newest one's on top of homepage", 
							"NA", "Alerts should be displayed in descending order with newest one's on top of homepage",
							strResult, strFlag);
				}
			} catch (Throwable e) {
				ReportFactory.reporterOutput(strTCID,"Verify Alerts are displayed in descending order with newest one's on top of homepage", "NA", "NA",
					e.getMessage().toString(), "Fail");
			}

	}
	
	
/*	
	public static void verifyAlertsShowingInDescendingOrderOnDateTime(String strAlertHeaderTextFromExcel,
			String strDepartmentNameFromExcel, String strTCID) throws Throwable {
		
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
							 * Get list of alerts and verify if sorted according to date
							 ***************/
				/*			if (ValidationFactory.isElementPresent(wbelPortletHeader)) {
								List<WebElement> numberOfAlerts = AlertHeaderTitle;
								LogFactory.info("Number of alerts are :" + numberOfAlerts.size());
								List<String> AlertsHeaderDates = GenericFactory.getListOfDatesByHeaderName(strAlertHeaderTextFromExcel);
								
								if (GenericFactory.verifyDateSortedOrder(AlertsHeaderDates) == true) {
									liMatchingDepartments.add(liDepartmentNameFromExcel.get(j));
									strResultPresentOnDepartment = strResultPresentOnDepartment.concat(liDepartmentNameFromExcel.get(j));
								} else {
									liNonMatchingDepartments.add(liDepartmentNameFromExcel.get(j));
									strResultNotPresentOnDepartment=strResultNotPresentOnDepartment.concat(liDepartmentNameFromExcel.get(j));
								}
							} 
							else {
									LogFactory.info("Alert portlet is not present on page:");
									strResultNotPresentOnDepartment=strResultNotPresentOnDepartment.concat(liDepartmentNameFromExcel.get(j));
									liNonMatchingDepartments.add(liDepartmentNameFromExcel.get(j));
							}
					}
				
					if(liMatchingDepartments.size()>0 ){   
						strFlag = "Pass";
						strResult="Alerts title dates are displayed in descending order on active departments :"+liMatchingDepartments;
					}
					else if(liMatchingDepartments.size()==0 || liNonMatchingDepartments.size()>0){
						strFlag = "Fail";
						strResult="Alerts title dates are NOT displayed in descending order on  departments  :"+liNonMatchingDepartments;
					}
					ReportFactory.reporterOutput(strTCID,"Verify alerts displayed in descending order with newest one's on top of department", strDepartmentNameFromExcel,"Alerts should be displayed in descending order with newest one's on top on active department pages", strResult, strFlag);
				}
			} 
			else {

					GenericFactory.navigateToHomePage();
					LogFactory.info("Verifying alerts displaying in descending order on date and time in homepage");
					/*********** * Get list of alerts and verify if sorted according to date ***************/
/*					List<WebElement> numberOfAlerts = BaseClass.wbDriver.findElements(By.xpath("//div[@class='section warning']"));
					List<String> AlertsHeaderDates = GenericFactory.getListOfDatesByHeaderName(strAlertHeaderTextFromExcel);
		
					if (GenericFactory.verifyDateSortedOrder(AlertsHeaderDates) == true) {
						strFlag = "Pass";
						strResult = "Alerts title dates are in descending order with newest one's on top by date on homepage";
					} 
					else {
							LogFactory.info("Alerts title dates are not in sorted order on homepage");
							strResult = "Alerts title dates are not in sorted order on homepage";
					}
					ReportFactory.reporterOutput(
							strTCID,"Verify alerts displayed in descending order with newest one's on top of homepage", 
							"NA", "Alerts should be displayed in descending order with newest one's on top of homepage",
							strResult, strFlag);
				}
			} catch (Exception e) {
				ReportFactory.reporterOutput(strTCID,"Verify alerts displayed in descending order with newest one's on top of homepage", "NA", "NA",
					e.getMessage().toString(), "Fail");
			}

	}
	
	*/
	
	/**************
	 * check alert header content format
	 *************************************************/

	/**
	 * @author Neeraja.mantri
	 * @createdAt 22-06-2018
	 * @param strTCID
	 * @Modified At 26-06-2018
	 * @throws Throwable
	 */
	public static void checkAlertHeaderContentWithFormat(String strTCID) throws Throwable {
		//By alertTitlesXpath = By.xpath("//div[@class='section warning']/div[2]/div//div/div[1]/span");
 
		
		String strFlag = "Fail";
		String strResult = "Alert Portlet is not available";

		/*****************
		 * check alert header title content format
		 ********************************************************/

		if (ValidationFactory.isElementPresent(wbelAlertPortlet)) {
			LogFactory.info("Alerts portlet is present");

			if (GenericFactory.headerTitleFormat(alertsTableTitle)) {
				strFlag = "Pass";
				strResult = "All Alerts header titles are displayed in the format --> Published Date ("+ BaseClass.dateformat + ") : Alert Title ";
			}
			else {
				strFlag = "Fail";
				strResult = "Invalid Hearder or Date format found in the Alerts Portlet ";
			}
			
			ReportFactory.reporterOutput(strTCID,"Verify alerts header titles on the Alert Portlet ",  "Data format is : " + BaseClass.dateformat,
					"Alerts header title should be in the format --> Published Date ("+ BaseClass.dateformat + ") : Alert Title ",
					strResult, strFlag);
			
		} else {
					ReportFactory.reporterOutput(strTCID,"Verify alerts header titles on the Alert Portlet ",  "Data format is : " + BaseClass.dateformat,
						"Alerts header title should be in the format --> Published Date ("+ BaseClass.dateformat + ") : Alert Title", strResult,
					strFlag);
		}

	}


	public static void verifyEmbededlinks(String strTCID) throws Throwable {
			
	        int respCode = 200;
	        List<String> emptyLinks = new ArrayList<String>();
	        List<String> brokenLinks = new ArrayList<String>();
	        List<String> CorrectLinks = new ArrayList<String>();
			List<WebElement> lstWebElement = GenericFactory.getLinksFromFrame(wbelAlertPortlet);
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
	                     ReportFactory.reporterOutput(strTCID,"Verify embeded links on Alert Portlet", "NA",
	                                   "Embeded links should not be broken or empty",
	                                   "Embeded links are not broken and working fine as expected :" + correctLinksString, "Pass");
	               } else {
	                     String brokenLinksString = String.join(",", brokenLinks);
	                     String stringEmptyLinks = String.join(",", emptyLinks);
	                     ReportFactory.reporterOutput(strTCID,"Verify Embeded Links on Alert Portlet","NA",
	                                   "Embeded links should not be broken or empty",
	                                   "Broken Links/Empty links are :" + brokenLinksString + " , " + stringEmptyLinks, "Fail");
	               }
	        } else {
	               System.out.println("No embeded links are present");
	               ReportFactory.reporterOutput(strTCID,"Verify Embeded Links on Alert Portlets.","NA",
	                            "Embeded links should not be broken or empty", "No links are present", "Pass");
	        }
	        
			}catch (Exception e) {
				LogFactory.error("e");
				String er = e.getMessage().toString().trim();
				ReportFactory.reporterOutput(strTCID,"Verify Embeded Links on Alert Portlets.", "NA","NA", er, "Fail");
				}
	 }
	
	

	public static boolean checkReadMoreAndCollapseLinkForAlert(String strTCID) throws Throwable {
		String strFlag = "Fail";
		String strResult = "Readmore link is not present";
		boolean booFlagReadmore = true;
		boolean booFlagCollapse = true;
		boolean readMoreLinkflag = true;

		try {
			List<WebElement> readMoreLink = wbelAlertPortlet
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
				strResult = "ReadMore link is displayed in Alerts portlet and working as expected";
			} else if (booFlagReadmore && !booFlagCollapse) {
				strFlag = "Fail";
				strResult = "ReadMore link is present for Alerts portlet but not working as expected";
			} else {
				strFlag = "Pass";
				strResult = "ReadMore link is not displayed for Alerts portlet";
			}
			ReportFactory.reporterOutput(strTCID,
					"Verify 'ReadMore & Collapse' links are displayed on Alerts portlet", "NA",
					"If 'ReadMore & Collapse' link is present then it should work as expected", strResult, strFlag);

		} catch (Exception e) {
				ReportFactory.reporterOutput(strTCID,
					"Verify 'ReadMore & Collapse' links are displayed on Alerts portlet", "NA","NA", e.getMessage(), "Fail");
		}
		return booFlagCollapse && booFlagReadmore;

	}
    

	public static void checkOrderOfPortletsOnDepartment(String strDepartmentNamesFromExcel,String strTCID ) throws Throwable 
	{
		String strFlag = "Fail";
		String strResult = "";
		String resultPortletOrder = "";
		boolean booAlertPresent=false;
		boolean booAnnouncementPresent=false;
		boolean booLinkPresent=false;
		boolean booFavoritePresent=false;
		List<String> liDepartmentNameFromExcel = GenericFactory.splitString(strDepartmentNamesFromExcel, ",");
		List<String> liMatchingDepartments=new ArrayList<String>();
		boolean booFlagCheckPortletOrder = false;
		try {
			if (liDepartmentNameFromExcel.size() > 0) {
				for (int j = 0; j < liDepartmentNameFromExcel.size(); j++) {
					WebElement department = GenericFactory.getDeptname(liDepartmentNameFromExcel.get(j));
					if (department != null) {
						department.click();
						LogFactory.info("Checking portlets order for department" + " : " + liDepartmentNameFromExcel.get(j));
						ArrayList<WebElement> listPortletframes = new ArrayList<WebElement>();
						if(isAlertPortletPresent().equalsIgnoreCase("Pass"))
						{
							listPortletframes.add(wbelAlertPortlet);
						booAlertPresent=true;
						
						}
						if(ValidationFactory.isElementPresent(Announcementportlet))
						{
						listPortletframes.add(Announcementportlet);
						booAnnouncementPresent=true;
						}
						if(ValidationFactory.isElementPresent(favoritesIndex))
						{
						listPortletframes.add(favoritesIndex);
						booFavoritePresent=true;
						}
						if(ValidationFactory.isElementPresent(Linksportlet))
						{
						listPortletframes.add(Linksportlet);
						booLinkPresent=true;
						}
						int intListPortletFrameSize = listPortletframes.size();
						
						switch (intListPortletFrameSize) {
						case 3:
							resultPortletOrder = "Portlets are not displayed as expected";
							liMatchingDepartments.add(liDepartmentNameFromExcel.get(j));
								if(listPortletframes.indexOf(wbelAlertPortlet)==0  )
								{
									LogFactory.info("Alert portlet is present at first and above other portlets");
									strFlag="Pass";
								}
								if(booAnnouncementPresent && booLinkPresent)
								{
								if(listPortletframes.indexOf(Announcementportlet)<listPortletframes.indexOf(linkPortletIndex))
								{
									LogFactory.info("Announcement portlet is present above link portlet");
									strFlag="Pass";
									resultPortletOrder = "Alerts portlet present above announcements & announcements portlet present above links portlet as expected on active departments"+liMatchingDepartments;
								}
								}
								else{
									LogFactory.info("portlets are not present in order");
									strFlag="Fail";
								}
							LogFactory.info("All the portlets are present in order");
							
							break;
						case 2:
							resultPortletOrder = "Portlets are not displayed as expected";
							liMatchingDepartments.add(liDepartmentNameFromExcel.get(j));
								if(booAlertPresent && booAnnouncementPresent)
									{
									if(listPortletframes.indexOf(Announcementportlet)>listPortletframes.indexOf(wbelAlertPortlet))
									LogFactory.info("Alert portlet present above announcement portlet");
									resultPortletOrder = "Alert portlet is displayed above announcements portlet as expected as expected on active departments"+liMatchingDepartments;


									
									strFlag="Pass";
									}
								else if(booAnnouncementPresent && booLinkPresent)
								{
									if(listPortletframes.indexOf(Announcementportlet)<listPortletframes.indexOf(Linksportlet))
										{
											LogFactory.info("Announcement portlet is present above Links portlet");
											strFlag="Pass";
											resultPortletOrder = "Announcement portlet is displayed above link portlet as expected on active departments"+liMatchingDepartments;
											 
										}
								}
								else if(booLinkPresent && booAlertPresent) {
								if(listPortletframes.indexOf(Linksportlet)>listPortletframes.indexOf(wbelAlertPortlet))
								{
									LogFactory.info("Alerts portlet and link portlet are present in order");
									resultPortletOrder = "Links portlet is displayed below alerts portlet as expected on active departments"+liMatchingDepartments;
									strFlag="Pass";
								}
								}
								else {
									LogFactory.info("Portlets are not present in order as expected");
									strFlag="Fail";
								}
							break;
						case 1:
								LogFactory.info("Only one portlet is present");
								liMatchingDepartments.add(liDepartmentNameFromExcel.get(j));
								if(ValidationFactory.isElementPresent(wbelAlertPortlet))
								{
								if(listPortletframes.indexOf(wbelAlertPortlet)==0 )
								{
									LogFactory.info("Only alert portlet is present");
									resultPortletOrder = "Only alert portlet is displayed as expected on active departments"+liMatchingDepartments;
									strFlag="Pass";
								}
								}
								else if(listPortletframes.indexOf(Linksportlet)==0)
								{
									LogFactory.info("Only link portlet is present");
									resultPortletOrder = "Only links portlet is displayed as expected on active departments"+liMatchingDepartments;
									strFlag="Pass";
									break;
								}
								else if(listPortletframes.indexOf(Announcementportlet)==0)
										{
											LogFactory.info("Only announcement portlet is present");
											resultPortletOrder = "Only announcment portlet is displayed as expected on active departments"+liMatchingDepartments;
											strFlag="Pass";
											break;
										}
							break;
						default:
							String resultOrder = "No portlets are displayed on active departments"+liMatchingDepartments;
							LogFactory.info("No portlets are present");
							strFlag="Fail";
						}
						
					}
					Homepage_POF.homepagepath.click();
					booAlertPresent=false;
					booAnnouncementPresent=false;
					booLinkPresent=false;
			}
				ReportFactory.reporterOutput("TC07_ALERTS_DEPARTMENT","Verify order of portlets displayed",strDepartmentNamesFromExcel,"Portlets should be displayed in order :Alert portlet in first,announcement portlet in second and links portlet in last",resultPortletOrder, strFlag);
			}
			else {
		
				LogFactory.info("No testdata is provided for departments in Excel sheet");
			}
		}
		
		catch (Throwable e) {
			LogFactory.info("Exception :" + e.getMessage().toString());
		}
	
	
}

	public static void checkOrderOfPortletsOnHomepage(String strDepartmentNamesFromExcel,String strTCID ) throws Throwable 
	{
		String strFlag = "Fail";
		String strResult = "";
		String resultPortletOrder = "";
		boolean booAlertPresent=false;
		boolean booAnnouncementPresent=false;
		boolean booFavoritePresent=false;
		
		List<String> liDepartmentNameFromExcel = GenericFactory.splitString(strDepartmentNamesFromExcel, ",");
		List<String> liMatchingDepartments=new ArrayList<String>();
		
		boolean booFlagCheckPortletOrder = false;
		try {
			if(!strDepartmentNamesFromExcel.equalsIgnoreCase("NA")) 
			{
			if (liDepartmentNameFromExcel.size() > 0) {
				LogFactory.info("Check portlet order for department");
				}
			}
			else {
				LogFactory.info(
						"Department is not present for checkOrderOfPortletsOnHomepage&Department method and is on homepage");
	
				ArrayList<WebElement> listPortletframes = new ArrayList<WebElement>();
						if(isAlertPortletPresent().equalsIgnoreCase("Pass"))
						{
							listPortletframes.add(wbelAlertPortlet);
						booAlertPresent=true;
						
						}
						if(ValidationFactory.isElementPresent(Announcementportlet))
						{
						listPortletframes.add(Announcementportlet);
						booAnnouncementPresent=true;
						}
						if(ValidationFactory.isElementPresent(favoritesIndex))
						{
						listPortletframes.add(favoritesIndex);
						booFavoritePresent=true;
						}
						
						int intListPortletFrameSize = listPortletframes.size();
				switch (intListPortletFrameSize) {
						case 3:
							resultPortletOrder = "Portlets are not displayed as expected";
							
								if(listPortletframes.indexOf(wbelAlertPortlet)==0  )
								{
									LogFactory.info("Alert portlet is present at first and above other portlets");
									strFlag="Pass";
								}
								if(booAnnouncementPresent && booFavoritePresent)
								{
								if(listPortletframes.indexOf(Announcementportlet)<listPortletframes.indexOf(favoritesIndex))
								{
									LogFactory.info("Announcement portlet is present above link portlet");
									strFlag="Pass";
									resultPortletOrder = "Alerts portlet present above announcements & announcements portlet present above favorites portlet as expected on homepage";
								}
								}
								else{
									LogFactory.info("portlets are not present in order");
									strFlag="Fail";
								}
							LogFactory.info("All the portlets are present in order");
							
							break;

						case 2:
							resultPortletOrder = "Portlets are not displayed as expected";
							
								if(booAlertPresent && booAnnouncementPresent)
									{
									if(listPortletframes.indexOf(Announcementportlet)>listPortletframes.indexOf(wbelAlertPortlet))
									LogFactory.info("Alert portlet present above announcement portlet");
									resultPortletOrder = "Alert portlet is displayed above announcements portlet as expected as expected on homepage";
									strFlag="Pass";
									}
								else if(booAnnouncementPresent && booFavoritePresent)
								{
									if(listPortletframes.indexOf(Announcementportlet)<listPortletframes.indexOf(favoritesIndex))
										{
											LogFactory.info("Announcement portlet is present above favorites portlet");
											strFlag="Pass";
											resultPortletOrder = "Announcement portlet is displayed above favorites portlet as expected on homepage";
										}
								}
								else if(booFavoritePresent && booAlertPresent) {
								if(listPortletframes.indexOf(favoritesIndex)>listPortletframes.indexOf(wbelAlertPortlet))
								{
									LogFactory.info("Alerts portlet and link portlet are present in order");
									resultPortletOrder = "Favorites portlet is displayed below alerts portlet as expected on homepage";
									strFlag="Pass";
								}
								}
								else {
									LogFactory.info("Portlets are not present in order as expected");
									strFlag="Fail";
								}
							break;
						case 1:
								LogFactory.info("Only one portlet is present");
						
								if(ValidationFactory.isElementPresent(wbelAlertPortlet))
								{
								if(listPortletframes.indexOf(wbelAlertPortlet)==0 )
								{
									LogFactory.info("Only alert portlet is present");
									resultPortletOrder = "Only alert portlet is displayed as expected on homepage";
									strFlag="Pass";
								}
								}
								else if(listPortletframes.indexOf(favoritesIndex)==0)
								{
									LogFactory.info("Only link portlet is present");
									resultPortletOrder = "Only favorites portlet is displayed as expected on homepage";
									strFlag="Pass";
									break;
								}
								else if(listPortletframes.indexOf(Announcementportlet)==0)
										{
											LogFactory.info("Only announcement portlet is present");
											resultPortletOrder = "Only announcment portlet is displayed as expected on homepage";
											strFlag="Pass";
											break;
										}
							break;
							default:
							String resultOrder = "No portlets are displayed on homepage";
							LogFactory.info("No portlets are present");
							strFlag="Fail";
						}
						
					Homepage_POF.homepagepath.click();
					booAlertPresent=false;
					booAnnouncementPresent=false;
					booFavoritePresent=false;
			
				ReportFactory.reporterOutput(strTCID,"Verify order of portlets displayed","NA","Portlets should be displayed in order :Alert portlet in first,announcement portlet in second and links portlet in last",resultPortletOrder, strFlag);
			}
			}
			catch (Throwable e) {
			LogFactory.info("Exception :" + e.getMessage().toString());
		}
		}


	public static void verifyAlertsOnTheAlertPortlet(String strWCMTCID, String userDefinedCountry, String wcmCountry,
			String userDefinedProducts, String wcmProducts, String contentType,String title) throws Throwable {

		String flag ="";
		String testCaseDescription = " Verify the Alert titles on the Alert portlet for content type : " + contentType;
		String inputValue = " <B>WCM Maping :</B> Country : " + wcmCountry + " ,Product Type : " + wcmProducts
				+  "<P>" + " <B>Dealer Maping :</B> Country : " + userDefinedCountry + " ,Product Type : " + userDefinedProducts;
		String expectedValue = "Alert portlet should show the valid alerts for the user";
		
		boolean booCountry = GenericFactory.userAndWCMCountryComparison(userDefinedCountry, wcmCountry);
		boolean booProduct = GenericFactory.userAndWCMProductTypeComparison(userDefinedProducts, wcmProducts);
		boolean booALertTitle = alertTitleComparison(title);
		String actualResult ="";

		if (booCountry && booProduct && booALertTitle) {
			actualResult = "User & WCM mapprings are matching so the Alert : " + title + " is appearing for the user";
			flag = "Pass";
		} else if (booCountry && !booProduct && !booALertTitle) {
			actualResult = "User Product Types are not matching with WCM mapping so the Alert : " + title + " is NOT appearing for the user";
			flag = "Pass";
		} else if (!booCountry && !booProduct && !booALertTitle) {
			actualResult = "User Country & Product Types are not matching with WCM mapping so the Alert : " + title + " is NOT appearing for the user";
			flag = "Pass";
		} else if (!booCountry && booProduct && booALertTitle) {
			actualResult = "User Country is not matching with WCM mapping and the Alert : " + title + " is appearing for the user";
			flag = "Fail";
		} else if (!booCountry && booProduct && !booALertTitle) {
			actualResult = "User Country is not matching with WCM mapping and the Alert : " + title + " is NOT appearing for the user";
			flag = "Pass";
		} else if (!booCountry && !booProduct && booALertTitle) {
			actualResult = "User Country & Product Types are not matching with WCM mapping but the Alert : " + title + " is appearing for the user";
			flag = "Fail";
		} else if (booCountry && booProduct && !booALertTitle) {
			actualResult = "User Country & Product Types are matching with WCM mapping but the Alert : " + title + " is NOT appearing for the user";
			flag = "Fail";
		} else if (booCountry && !booProduct && booALertTitle) {
			actualResult = "User Product Types are NOT matching with WCM mapping but the Alert : " + title + " is appearing for the user";
			flag = "Fail";
		} else {
			actualResult = "Invalid Content";
			flag = "Fail";
		}
			ReportFactory.reporterOutput(strWCMTCID, testCaseDescription, inputValue,expectedValue, actualResult, flag);
	}

}

