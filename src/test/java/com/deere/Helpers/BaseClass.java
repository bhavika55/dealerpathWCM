	/**
	 * This class is use to initiate the webElements of all mentioned pages
	 * Project    : DealerPath
	 * @author shrishail.baddi
	 * @createdAt 15-05-2018
	 * 
	 * @ModifiedBy shrey.choudhary
	 * @modifiedAt 07-06-2018
	 */

package com.deere.Helpers;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.support.PageFactory;

import com.deere.PageFactory.Alerts_POF;
import com.deere.PageFactory.Announcements_POF;
import com.deere.PageFactory.Favourites_POF;
import com.deere.PageFactory.Homepage_POF;
import com.deere.PageFactory.Login_Page_POF;
import com.deere.PageFactory.PortalLeftNavigation_POF;
import com.deere.PageFactory.PortletLinksPage_POF;
import com.deere.PageFactory.ProductSegment_POF;
import com.deere.PageFactory.UtilityLinks_POF;
import com.deers.alerts_WCM.Alert_WCM_POF;

/*import com.deere.PageFactory.ProductSegment_POF;
import com.deere.PageFactory.Search_POF;
import com.deere.PageFactory.UtilityLinks_POF;
import com.deere.PageFactory.Alerts_POF;
import com.deere.PageFactory.Announcements_POF;
*/

public class BaseClass {
	
	//Global/Environment Variables & Constants
	public static WebDriver wbDriver = null;
	public static String strWorkingDir = System.getProperty("user.dir").toString();
	public static String URL="https://dealerpathqual.tal.deere.com/";
	public static String strBrowserType="EDGE";
	public static String strUserName="ss88593";
	public static String strPassword="dfascr65";
	public static String ENABLE_LOGS = "ON";  //  possible values "ON"  or "OFF"
	public static String strUserPrefLang="ENGLISH";
	public static String strUserRole="Dealer";
	public static String strSiteRegion = "R4_AUNZ";
	public static String strDataPath = strWorkingDir + "\\TestData\\testdatasample.xlsx";
	public static String strDataOutputPath = strWorkingDir + "\\ExportToExcel\\ExportToExcel";
	public static String wcmDataOutputPath = strWorkingDir + "\\WCMContent\\WCMContent";
	public static String dateformat=  "dd-mm-yyyy"; // "yyyy-mm-dd";
	public static String strUserRACFID= "";
	public static Properties propConfig = null;
	public static String strTestCase="Index TC";
	public static String strScreenshotDir;
	public static String ENABLE_SCREENSHOT;
	public static boolean errorUserFound = false;
	
	public static String alertRegion="R4";
	@SuppressWarnings("rawtypes")
	public static List<LinkedHashMap> wcmExceldata;

/*	// Constants

	public final String R3_HIS = "R3 Hispano";
	public final String JDIN= "R2 JDIN";
	public final String R1_SS= "R1 Sub Saharan Africa"; 
	public final String R1_AS="R1 JD Asia"; 
	public final String R4_USA  = "R4 USA/Canada";
	public final String R1_CH="R1 China"; 
	public final String R1_TH="R1 Thailand";
	public final String R1_IN="R1 India"; 
	public final String R4_AUNZ= "R4 Australia/New Zealand";
	public final String R3_Bra = "R3 Brasil"; */
	
	public static List<String> headerList;
	public static Map<String,String> excelList;
	public static List<Map<String,String>> finalResultforExcel = new ArrayList<>();
	//pages objects
	public static Login_Page_POF loginPageFactory;
	public static Homepage_POF homePageFactory;
	//public static Homepage_POF homePageFactory1;
	public static PortletLinksPage_POF portletLinksFactory;
	public static PortalLeftNavigation_POF  portalLeftNavFactory;
	public static HashMap <String,String> mapAddtionalTestcase;
	public static String flagDealerType;
	public static String flagDealerPrinciple;
	public static String flagAddtionalTestcases;
	
	public static List<LinkedHashMap> userWCMData;
	public static Alerts_POF alertPageFactory;
	public static Announcements_POF announcementFactory;
	public static UtilityLinks_POF  utilityLinksFactory;
	public static Favourites_POF  favouritesFactory;
	//public static Search_POF searchFactory;
	public static ProductSegment_POF productSegmentFactory;
	
	public static Alert_WCM_POF alertContent;

	
	
	/**
	 * This method is use to initiate the webElements of all mentioned pages
	 * @author shrishail.baddi
	 * @createdAt 15-05-2018
	 * 
	 * @ModifiedBy shrey.choudhary
	 * @modifiedAt 07-06-2018
	 */
	public static void initPageElements() {
		
		loginPageFactory=PageFactory.initElements(BaseClass.wbDriver, Login_Page_POF.class);
		//	homePageFactory=PageFactory.initElements(BaseClass.wbDriver, Homepage_POF_Old.class);
		homePageFactory=PageFactory.initElements(BaseClass.wbDriver, Homepage_POF.class);
		portalLeftNavFactory=PageFactory.initElements(BaseClass.wbDriver, PortalLeftNavigation_POF.class);
		portletLinksFactory= PageFactory.initElements(BaseClass.wbDriver, PortletLinksPage_POF.class);
		alertPageFactory=PageFactory.initElements(BaseClass.wbDriver, Alerts_POF.class);
		announcementFactory=PageFactory.initElements(BaseClass.wbDriver, Announcements_POF.class);
		utilityLinksFactory=PageFactory.initElements(BaseClass.wbDriver, UtilityLinks_POF.class);
		favouritesFactory=PageFactory.initElements(BaseClass.wbDriver, Favourites_POF.class);
		//searchFactory= PageFactory.initElements(BaseClass.wbDriver, Search_POF.class);
		productSegmentFactory= PageFactory.initElements(BaseClass.wbDriver, ProductSegment_POF.class);
	}
	
	public static String getExcelDataByTestCaseID(String strTestId) throws Exception{
		 HashMap<String, String> Data = mapAddtionalTestcase;
		 if(Data.containsKey(strTestId)){
			 String value = Data.get(strTestId);
//			 prefferedLanguage = Data.get
			 return value;
		 
		 	}
		 	else{
		 		LogFactory.error("------ TestCase_ID :  "+strTestId+" is not Avaliable ");
		 		return "None";
		 	}
		}
}
