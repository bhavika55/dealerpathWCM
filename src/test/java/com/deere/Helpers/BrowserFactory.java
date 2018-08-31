
/* 
 * Project    : DealerPath
 * Script     : BrowserFactory
 * Author     : Shrishail Baddi
 * Date       : April.02.2018
 * Last Modified On:
 * Modified By :
 */




package com.deere.Helpers;

import java.util.HashMap;
import java.util.Map;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxOptions;
import org.openqa.selenium.firefox.FirefoxProfile;
import org.openqa.selenium.htmlunit.HtmlUnitDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.events.EventFiringWebDriver;

import com.deere.Helpers.BaseClass;

public class BrowserFactory {
	
	
	public static EventFiringWebDriver driver = null;
	
	private static String strChromeDriverPath = BaseClass.strWorkingDir + "\\src\\test\\resources\\Drivers\\chromedriver1.exe";
	private static String strIEDriverPath = BaseClass.strWorkingDir+ "\\src\\test\\resources\\Drivers\\IEDriverServer_3.11.32.exe";
	private static String strGeckoDriverPath = BaseClass.strWorkingDir + "\\src\\test\\resources\\Drivers\\geckodriver1.exe";
	private static String strEdgeDriverPath = BaseClass.strWorkingDir + "\\src\\test\\resources\\Drivers\\EDGEWebDriver_15063.exe";
	
	private static String strBrowser = BaseClass.strBrowserType;			

	
	
	
	
	/*public static DesiredCapabilities Chrome_Profile() throws Exception{
		String FilePath = createDirectory();
		
		DesiredCapabilities capabilities = new DesiredCapabilities().chrome();
		
		ChromeOptions options = new ChromeOptions();
	    Map<String, Object> prefs = new HashMap<String, Object>();
	    prefs.put("profile.default_content_settings.popups", 0);
	    prefs.put("download.default_directory", FilePath);
	    options.setExperimentalOption("prefs", prefs);
	    options.addArguments("--test-type");
	    capabilities.setCapability(ChromeOptions.CAPABILITY, options);
	    return capabilities;
	    
	}*/
	
	
	
	
	/*public static DesiredCapabilities FF_Profile() throws Exception
	{
		DesiredCapabilities capabilities = new DesiredCapabilities().firefox();
		
		String FilePath = createDirectory();
		
		FirefoxProfile profile = new FirefoxProfile();
		profile.setPreference("browser.download.folderList", 2);
		profile.setPreference("browser.download.manager.showWhenStarting", false);
		profile.setPreference("browser.helperApps.alwaysAsk.force", false);
		profile.setPreference("browser.download.manager.alertOnEXEOpen", false);
		profile.setPreference("browser.download.manager.focusWhenStarting", false);
		profile.setPreference("browser.download.manager.useWindow", false);
		profile.setPreference("browser.download.manager.showAlertOnComplete", false);
		profile.setPreference("browser.download.manager.closeWhenDone", false);
		profile.setPreference("browser.download.dir", FilePath);
		profile.setPreference("services.sync.prefs.sync.browser.download.manager.showWhenStarting", false);
		profile.setPreference("pdfjs.disabled", true);
		profile.setPreference("browser.helperApps.neverAsk.saveToDisk", "text/csv,application/x-msexcel,application/excel,application/x-excel,application/vnd.ms-excel,image/png,image/jpeg,text/html,text/plain,application/msword,application/xml,application/pdf");
		
		capabilities.setCapability(FirefoxDriver.PROFILE, profile);
		return capabilities;
	}
	*/
	
	
	
	
	
	
	
	
	
	
	
			
	/*This method is used to launch the browser - with supplied URL &  Browser type
	 * 
	 * @exception Exception: If any exception is found	
	 */
	
	@SuppressWarnings("deprecation")
	public static void initiateDriver(){
		
		
		try {
				
				if (!strBrowser.isEmpty() && !strBrowser.equals(null)) {
					
					// checking the type of browser
					System.out.println("Checking Browser");
					if(strBrowser.equalsIgnoreCase("Chrome")){
						
					System.setProperty("webdriver.chrome.driver", strChromeDriverPath );
						BaseClass.wbDriver = new ChromeDriver();
					
						
					/*	
						
						DesiredCapabilities capabilities = new DesiredCapabilities().chrome();
						
						ChromeOptions options = new ChromeOptions();
					    Map<String, Object> prefs = new HashMap<String, Object>();
					    prefs.put("profile.default_content_settings.popups", 0);
					   // prefs.put("download.default_directory", FilePath);
					    options.setExperimentalOption("prefs", prefs);
					    options.addArguments("--test-type");
					    capabilities.setCapability(ChromeOptions.CAPABILITY, options);
					    BaseClass.wbDriver = new ChromeDriver();
				    
					    */
					    
						/*System.setProperty("webdriver.chrome.driver", strChromeDriverPath );
					    ChromeOptions options = new ChromeOptions();
					    				    
					    BaseClass.wbDriver = new ChromeDriver(options);
					    */
						
					}else if(strBrowser.equalsIgnoreCase("Firefox")){
						
						System.setProperty("webdriver.gecko.driver", strGeckoDriverPath);
						
						/*DesiredCapabilities capabilities = DesiredCapabilities.firefox();
						capabilities.setCapability("marionette", true);
						BaseClass.wbDriver = new FirefoxDriver(capabilities);
						*/
						
						
						/*DesiredCapabilities capabilities = new DesiredCapabilities().firefox();
						FirefoxProfile profile = new FirefoxProfile();
						profile.setPreference("browser.download.folderList", 2);
						profile.setPreference("browser.download.manager.showWhenStarting", false);
						profile.setPreference("browser.helperApps.alwaysAsk.force", false);
						profile.setPreference("browser.download.manager.alertOnEXEOpen", false);
						profile.setPreference("browser.download.manager.focusWhenStarting", false);
						profile.setPreference("browser.download.manager.useWindow", false);
						profile.setPreference("browser.download.manager.showAlertOnComplete", false);
						profile.setPreference("browser.download.manager.closeWhenDone", false);
						//profile.setPreference("browser.download.dir", FilePath);
						profile.setPreference("services.sync.prefs.sync.browser.download.manager.showWhenStarting", false);
						profile.setPreference("pdfjs.disabled", true);
						profile.setPreference("browser.helperApps.neverAsk.saveToDisk", "text/csv,application/x-msexcel,application/excel,application/x-excel,application/vnd.ms-excel,image/png,image/jpeg,text/html,text/plain,application/msword,application/xml,application/pdf");
						
						capabilities.setCapability(FirefoxDriver.PROFILE, profile);
						BaseClass.wbDriver = new FirefoxDriver(capabilities);
					*/	
						
						
						FirefoxOptions options = new FirefoxOptions();
						
						options.setBinary("C:\\Program Files (x86)\\Mozilla Firefox\\firefox.exe");
						FirefoxDriver driver = new FirefoxDriver(options);
						BaseClass.wbDriver  = new FirefoxDriver(options);

				        
				        
				    }
						
						
					
				else if(strBrowser.equalsIgnoreCase("IE")){
						
						System.setProperty("webdriver.ie.driver", strIEDriverPath);
						
						
						DesiredCapabilities capabilities = DesiredCapabilities.internetExplorer();
						capabilities.setCapability(InternetExplorerDriver.INTRODUCE_FLAKINESS_BY_IGNORING_SECURITY_DOMAINS,true);
						capabilities.setCapability("requireWindowFocus", true);
						capabilities.setCapability("ignoreZoomSetting", true);
						
						
						/*capabilities.setCapability(InternetExplorerDriver.NATIVE_EVENTS, false);
						capabilities.setCapability(InternetExplorerDriver.IGNORE_ZOOM_SETTING, true);
						capabilities.setCapability("allow-blocked-content", true);
						capabilities.setCapability("allowBlockedContent", true);
						
						*/
						
						BaseClass.wbDriver = new InternetExplorerDriver(capabilities);

						
					}else if(strBrowser.equalsIgnoreCase("EDGE")){
						
						System.setProperty("webdriver.edge.driver", strEdgeDriverPath);
						BaseClass.wbDriver  = new EdgeDriver();

						
					}else if(strBrowser.equalsIgnoreCase("HtmlUnit")){
						
				/*		DesiredCapabilities caps = new DesiredCapabilities();
						caps.setJavascriptEnabled(true); // not really needed: JS enabled by default
						caps.setCapability(PhantomJSDriverService.PHANTOMJS_EXECUTABLE_PATH_PROPERTY, "C:/phantomjs.exe");
						WebDriver driver = new PhantomJSDriver(caps);*/
						
					}
					
					driver = new EventFiringWebDriver(BaseClass.wbDriver);
					
					// putting an implicit wait after every Action or Event
					driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
					
					// opening the browser
					LogFactory.info ("opening the browser....");
					driver.navigate().to(BaseClass.URL);
					driver.manage().window().maximize();
					driver.manage().deleteAllCookies();
				}
				else
				{
					
					LogFactory.info ("Invalid Browser");
				}
		
		}
		
		catch(Exception e){
			System.out.println (e.getMessage());
		}
}
	
	
	
	//Method to close the driver
	
	public static void closeDriver() {
		if(BaseClass.wbDriver!=null) {
			driver.close();
			BaseClass.wbDriver.close();
			
		}
		
	}
	
}