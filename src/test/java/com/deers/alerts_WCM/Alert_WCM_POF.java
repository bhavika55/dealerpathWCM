package com.deers.alerts_WCM;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.How;
import com.deere.Helpers.BaseClass;
import com.deere.Helpers.ValidationFactory;
import com.deere.Helpers.WaitFactory;

public class Alert_WCM_POF extends BaseClass{

	static WebDriver alrtDriver;
	private static String filename = "";
	static XSSFWorkbook workbook = null;
	static XSSFSheet spreadsheet = null;
	static String alertName=null;
	
	private static XSSFWorkbook wcmbook;
	private static XSSFSheet wcmdataSheet;
	
	static int testcaseNumber=1;
	
	static String testCaseID="WCM_TC";
	public static List<Map<String,String>> finalResultforExcel = new ArrayList<>();
	
	
	public Alert_WCM_POF(WebDriver driver)
	{
		this.alrtDriver=driver;
		
	}
	
	
	
	@FindBy(how = How.XPATH, using = "//a[contains(@id,'pageSizeFiftywcmTable')]")
	public static WebElement bottomPagenumber;
	
	@FindBy(how = How.XPATH, using = "//table//tr[contains(@id,'wcmTable_')]")
	public static WebElement allAlerts;
	
	
	@FindBy(how = How.XPATH, using = "//td[contains(.,'Published') and @class='lotusMeta']")
	public static WebElement allPublishedAlerts;
	
	
	@FindBy(how = How.XPATH, using = "h4[@role='presentation']//a[contains(.,'Content')]")
	public static WebElement contentSection;

	
	@FindBy(how = How.XPATH, using = "h4[@role='presentation']//a[contains(.,'Alerts')]")
	public static WebElement alertsSection;
	
	
	@FindBy(how = How.XPATH, using = "//table//tr[contains(@id,'wcmTable_')]")
	public static List<WebElement> totalAlerts;	
	
	
	@FindBy(how = How.XPATH, using = "//td[contains(.,'Published') and @class='lotusMeta']")
	public static List<WebElement> totalPublishedAlerts;

	public static void clickOnBottomNumber() throws Throwable
	{
		
		try
		{
			Thread.sleep(3000);
			
			WebElement pageNumber = null;

			pageNumber = ValidationFactory.getElementIfPresent(By.xpath("//a[contains(@id,'pageSizeFiftywcmTable')]"));
			
			if(pageNumber != null)
			{
				
				pageNumber.click();
				System.out.println("Page Number link clicked sucessfully");
				
			}
			
			
		}
		catch(Exception e)
		{
			e.printStackTrace();
			System.out.println("Error while clicking the expand number link:"+e.getMessage().toString());
		}
		
	}
	
		
	public static void readWCMAlertContent() throws Throwable
	{
				
		System.out.println("**Fetching individual alerts content**");		
		try
		{
		
			
			moveInsideWCMContents("Alerts");
			
			System.out.println("***All alerts data fetched***");
			System.out.println("*Now navigating to Announcement section*");
			BaseClass.wbDriver.findElement(By.xpath("//li[@class='wcmBreadcrumbsElement']//a[contains(.,'Content')]")).click();
			
			BaseClass.wbDriver.findElement(By.xpath("//h4[@role='presentation']//a[contains(.,'My DealerPath')]")).click();
			
			BaseClass.wbDriver.findElement(By.xpath("//h4[@role='presentation']//a[contains(.,'Announcements')]")).click();
			
			
			Thread.sleep(3000);
		
			moveInsideWCMContents("Announcement");
			
		}
		
	

	catch(Exception e)
	{
		System.out.println("Link not clicked "+e.getMessage().toString());
	}
		
	
	}
	
	
	
	
	public static void moveInsideWCMContents(String wcmsection) throws Throwable
	{
		
		
	 try {
		 
		 System.out.println("***fetching WCM contents for "+wcmsection+" ***");
		 List<WebElement> publishedAlerts=ValidationFactory.getElementsIfPresent(By.xpath("//td[contains(.,'Published') and @class='lotusMeta']/preceding-sibling::td[1]//a"));
			
			Iterator<WebElement> iter = publishedAlerts.iterator();
			
			
			ArrayList<String> alertsLList = new ArrayList<String>();
			
			while(iter.hasNext()) {
				WebElement w = iter.next();
				
				String alertName=w.getText();
				alertsLList.add(alertName);
		 
			}
				System.out.println("Total Published "+wcmsection+" in list are:"+alertsLList.size());
							 
		
			 for(int i=0;i<alertsLList.size();i++) {
				 
				 String alert=alertsLList.get(i);
			    	
				 System.out.println("Fetching content for "+wcmsection+" :"+alert);
				
				 WebElement alert1=BaseClass.wbDriver.findElement(By.xpath("//a[contains(.,'"+alert+"')]"));
				 
				 
				 	alert1.click();
				 	
			    	String wcmTCID=testCaseID+testcaseNumber;
			    	
			    	   
			    	   writeWCMToExcel(wcmsection,wcmTCID);
			      
			        		        
			    	   writeWCMHeaderContentFinalToExcel();
			        
			        testcaseNumber++;
			       
			        String alertClose="//a[@id='close_controllable']";
			        WebElement close= WaitFactory.explicitWaitByXpath(alertClose);
			        if(close !=null)
			        {		        
			        	close.sendKeys(Keys.RETURN);	
			        	System.out.println("**Closed the "+wcmsection+" ** ");
			        }
			        	
	 										
			 }
	 }
			 catch(Exception e)
			 {
				 
				 System.out.println("Error while writing contents for "+ wcmsection+" " +e.getMessage().toString());
			 }
	 }
		
		
	
	
	
	public static void moveToAlertSection(WebElement alertRegionLanguage) throws Throwable
	{
					
	try
	{
		System.out.println("**Inside Alert navigation method**");
		
		//clickOnBottomNumber();
		
		
		if (alertRegionLanguage != null) {

			alertRegionLanguage.click();
			
			System.out.println("***clciking content section");
			
			ValidationFactory.getElementIfPresent(By.xpath("//h4[@role='presentation']//a[contains(.,'Content')]")).click();
			
			System.out.println("***clciking Alerts section");
			
			BaseClass.wbDriver.findElement(By.xpath("//h4[@role='presentation']//a[contains(.,'Alerts')]")).click();
		
		}
		
			clickOnBottomNumber();
			
	
		List<WebElement> totalAlerts = null;

		totalAlerts = ValidationFactory.getElementsIfPresent(By.xpath("//table//tr[contains(@id,'wcmTable_')]"));

		if(totalAlerts != null)
		{
			
			System.out.println("Total Alerts available are: "+totalAlerts.size());
		}
				
		List<WebElement> totalPublishedAlerts = null;

		totalPublishedAlerts = ValidationFactory.getElementsIfPresent(By.xpath("//td[contains(.,'Published') and @class='lotusMeta']"));

		if(totalPublishedAlerts != null)
		{
			
			System.out.println("Total published alerts are: "+totalPublishedAlerts.size());
		}
		
		
				
		}
	catch(Exception e)
	{
		e.printStackTrace();
		System.out.println("Couldn't navigate to alert section "+e.getMessage().toString());
	}
		
	
	}
	
	
	
	public static void applyfilter() throws Throwable{
		
		try {
			System.out.println("***Applying filter***");
			
			
			BaseClass.wbDriver.findElement(By.xpath("//a[contains(.,'Filter')]")).click();
			
			
			String clickFilter="//*[@id='ibm_wcm_widget_filter_FilterField_0_menuLink']";
			WebElement filterclicked= WaitFactory.explicitWaitByXpath(clickFilter);
			
			filterclicked.click();
			
			
			//BaseClass.wbDriver.findElement(By.xpath("//*[@id='ibm_wcm_widget_filter_FilterField_0_menuLink']")).sendKeys(Keys.RETURN);
			
			String selectingStatus="//td[@id='ibm_wcm_widget_filter_FilterField_0WORKFLOW_STATUS_text']";
			WebElement statusSelect= WaitFactory.explicitWaitByXpath(selectingStatus);
			
			statusSelect.click();
			
			//BaseClass.wbDriver.findElement(By.xpath("//td[@id='ibm_wcm_widget_filter_FilterField_0WORKFLOW_STATUS_text']")).click();		
			
			BaseClass.wbDriver.findElement(By.xpath("//*[@id='W1657b2fb178OkBtn']")).click();
			
		}
		catch(Exception e)
		{
				System.out.println("Unable to apply filter "+e.getMessage().toString());
		}
		
	}
	
	public static void moveToAnnouncementSection(String alertRegionLanguage) throws Throwable
	{
					
	try
	{
		System.out.println("**Inside Anouncement navigation method**");
		
		Thread.sleep(3000);

		
		clickOnBottomNumber();
		
		Thread.sleep(3000);
		
		WebElement alertSection = null;

		alertSection = ValidationFactory.getElementIfPresent(By.xpath("//a[contains(.,'"+alertRegionLanguage+"')]"));
		


		if (alertSection != null) {

			alertSection.click();
			
			System.out.println("***clciking content section");
			
			ValidationFactory.getElementIfPresent(By.xpath("//h4[@role='presentation']//a[contains(.,'Content')]")).click();
			
			System.out.println("***clciking MydealerPath section");
			
			ValidationFactory.getElementIfPresent(By.xpath("//h4[@role='presentation']//a[contains(.,'My DealerPath')]")).click();
			
			ValidationFactory.getElementIfPresent(By.xpath("//h4[@role='presentation']//a[contains(.,'Announcements')]")).click();
			
							}
		
		}
	
	catch(Exception e)
	{
		System.out.println("Link not clicked "+e.getMessage().toString());
	}
		
	
	}


	public static List<WebElement> identifyAlllanguages(String alertRegion)throws Throwable {

		
		try {
			
			
			clickOnBottomNumber();
			
			Thread.sleep(5000);
			List<WebElement> allLanguages= ValidationFactory.getElementsIfPresent(By.xpath("//a[contains(.,'"+alertRegion+"_CONTENT_')]"));
			
			
			System.out.println("***Total Languages available for the Region "+alertRegion+" are "+allLanguages.size());
			
			Iterator<WebElement> iter = allLanguages.iterator();

					
			while(iter.hasNext()) {
			    
					WebElement we = iter.next();

			    	System.out.println("**Region available is: "+we.getText());
			         
						}
			return allLanguages;
			
			}
		
		
		catch(Exception e)
		{
			
			System.out.println("elements not found"+e.getMessage().toString());
			
		}
		return null;
	}
	
	
	public static void writeWCMHeaderContentFinalToExcel() throws Throwable
	{
		
		try
		{
			System.out.println("***Writing final content into WCM Excel***");
			writeWCMHeader(filename, BaseClass.headerList);

			writeWCMRow(filename, BaseClass.finalResultforExcel);

		}
		catch(Exception e)
		{
			
			System.out.println("error while writing WCM content excel"+e.getMessage().toString());
		}
	}

	
	public static void createWCMExcel() throws Throwable{
		
				try {

			DateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy_HH-mm-ss");
			Date date = new Date();
			filename = wcmDataOutputPath + dateFormat.format(date) + ".xlsx";

			System.out.println("**WCM Excel created successsfully**");
			
	System.out.println("**Adding WCM content Headers into List**");
			 BaseClass.headerList= new  ArrayList<String>();
			   
			  BaseClass.headerList.add("Test Case ID"); 
			  BaseClass.headerList.add("EXECUTE");
			  BaseClass.headerList.add("URL");
			  BaseClass.headerList.add("Library");
			  BaseClass.headerList.add("Multilingual");
			  BaseClass.headerList.add("DepartmentName");
			  BaseClass.headerList.add("2ndLevel");
			  BaseClass.headerList.add("3rdLevelIndexPage");
			  BaseClass.headerList.add("3rdLevelIndexPageCategories");
			  BaseClass.headerList.add("3rdLevelIndexPageNestedCategories");
			  BaseClass.headerList.add("3rdLevelLandingPage");
			  BaseClass.headerList.add("3rdLevelChildIndexPage");
			  BaseClass.headerList.add("3rdLevelChildIndexPageCategories");
			  BaseClass.headerList.add("3rdLevelChildIndexPageNestedCategories");
			  BaseClass.headerList.add("3rdLevelGrandChildIndexPage");
			  BaseClass.headerList.add("3rdLevelGrandChildIndexPageCategories");
			  BaseClass.headerList.add("3rdLevelGrandChildIndexPageNestedCategories");
			  BaseClass.headerList.add("3rdLevelFolder");
			  BaseClass.headerList.add("4thLevelIndexPage");
			  BaseClass.headerList.add("4thLevelIndexPageCategories");
			  BaseClass.headerList.add("4thLevelIndexPageNestedCategories");
			  BaseClass.headerList.add("4thLevelLandingPage");
			  BaseClass.headerList.add("4thLevelChildIndexPage");
			  BaseClass.headerList.add("4thLevelChildIndexPageCategories");
			  BaseClass.headerList.add("4thLevelChildIndexPageNestedCategories");
			  BaseClass.headerList.add("4thLevelGrandChildIndexPage");
			  BaseClass.headerList.add("4thLevelGrandChildIndexPageCategories");
			  BaseClass.headerList.add("4thLevelGrandChildIndexPageNestedCategories");
			  BaseClass.headerList.add("ContentType");
			  BaseClass.headerList.add("IndexPageContentType");
			  BaseClass.headerList.add("Title");
			  BaseClass.headerList.add("Keywords");
			  
			  BaseClass.headerList.add("DocPath");
			  BaseClass.headerList.add("Link");
			  BaseClass.headerList.add("Description");
			  BaseClass.headerList.add("ReleaseDate");
			  BaseClass.headerList.add("Column4");   
			  BaseClass.headerList.add("Column5");
			  BaseClass.headerList.add("MRU-Country");
			  BaseClass.headerList.add("ProductType");
			  BaseClass.headerList.add("DealerType (Main/Sub)");
			  BaseClass.headerList.add("Index_Page_Template");
			  BaseClass.headerList.add("RACFGroups");
			  BaseClass.headerList.add("CopyToDepartment");
			  BaseClass.headerList.add("Comments");
			 
					      
		}
		catch(Exception e)
		{
			
			System.out.println("error while creating WCM content excel"+e.getMessage().toString());
		}
		
	}
	
	
	
	public static void writeWCMToExcel(String wcmsection,String wcmTCID) throws Throwable{
		String contentType;
		
		try {

			
			

contentType=BaseClass.wbDriver.findElement(By.xpath("//*[@id='content_template']")).getText();
String[] cType=contentType.split("/");
String conType =  cType[cType.length-1].trim();

			
String title=BaseClass.wbDriver.findElement(By.xpath("//*[@id='id_ctrl_titlecom.aptrix.pluto.content.Content']")).getText();

String location=BaseClass.wbDriver.findElement(By.xpath("//*[@id='locationcom.aptrix.pluto.content.Content']")).getText();

/*

String country=BaseClass.wbDriver.findElement(By.xpath("//label[.='MRU-Country']/following::div[1]")).getText();

Actions actions = new Actions(BaseClass.wbDriver);
actions.moveToElement(country1);
actions.perform();


((JavascriptExecutor) wbDriver).executeScript("arguments[0].scrollIntoView(true);", country1);
Thread.sleep(500);

String productType=BaseClass.wbDriver.findElement(By.xpath("//label[.='Product Type']/following::div[1]")).getText();

*/

			if(wcmsection.equals("Alerts"))
			{
				String library=BaseClass.wbDriver.findElement(By.xpath("//*[@id='breadcrumb_library']")).getText();			
			

				excelOutput(wcmTCID,library,conType,title,"USA",location,"Agriculture");
			}
			
			
			else if(wcmsection.equals("Announcement"))
			{
				String department=BaseClass.wbDriver.findElement(By.xpath("//label[.='Department']/following::div[1]")).getText();
				String copyToDepartment=BaseClass.wbDriver.findElement(By.xpath("//label[.='Copy Department']/following::div[1]")).getText();
				
				excelOutput(wcmTCID,department,contentType,title,"USA","Agriculture",copyToDepartment);
			}
			
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			System.out.println("Error while fetching content from "+wcmsection+" "+e.getMessage().toString());


		}
	}
	
		
	
	
	
	public static void excelOutput(String wcmTCID,String library,String contentType,String strTitle,String strCountry,
			String strLocation, String strProductType) throws Throwable {
		
		
		System.out.println("***mapping alerts contents into List***");
		try {

		 
		  BaseClass.excelList = new LinkedHashMap<String,String>();
		  
		  BaseClass.excelList.put("Test Case ID", wcmTCID);
		  BaseClass.excelList.put("Library", library);
		  BaseClass.excelList.put("ContentType", contentType);
		  BaseClass.excelList.put("Title", strTitle);		  
		  BaseClass.excelList.put("MRU-Country", strCountry);
		 // BaseClass.excelList.put("Location", strLocation);
		  BaseClass.excelList.put("ProductType", strProductType);
		  
		     System.out.println("Key Value hashmap: "+BaseClass.excelList);
		  BaseClass.finalResultforExcel.add(BaseClass.excelList);
		}
		catch(Exception e) {
		System.out.println("Error while mapping content to list "+e.getMessage().toString());
		}

	}
	
	
	
	
	
	public static String writeWCMHeader(String fileName, List<String> headerList) throws IOException {
		FileOutputStream fos = new FileOutputStream(new File(fileName));
		XSSFWorkbook book = new XSSFWorkbook();
		
		XSSFSheet sheet;
		
		sheet = book.createSheet("WCM Content");
				
		Row row = sheet.createRow(0);

		int cellNumber = 0;
		Font font = book.createFont();
		//font.setBold(true);
		font.setFontHeightInPoints((short) 9);
		font.setColor(IndexedColors.DARK_YELLOW.getIndex());
		font.setBold(true);
		
		CellStyle cellStyle1 = book.createCellStyle();

		for (String header : headerList) {
			Cell cell = row.createCell(cellNumber++);
			cellStyle1.setFont(font);
			/*		cellStyle1.setFillForegroundColor(IndexedColors.BRIGHT_GREEN.getIndex());*/
			
			cellStyle1.setFillForegroundColor(IndexedColors.GREEN.getIndex());
			
			cellStyle1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			//cellStyle1.setAlignment(CellStyle.ALIGN_CENTER);
			//cellStyle1.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
			
			cell.setCellStyle(cellStyle1);
			cell.setCellValue(header);
			sheet.autoSizeColumn(cellNumber);

		}
		book.write(fos);
		book.close();
		fos.close();
		return fileName;
	}
	
	
	public static void writeWCMRow(String fileName, List<Map<String, String>> rowList)
			throws IOException, InvalidFormatException {
		File oFile = new File(fileName);
		FileInputStream input = new FileInputStream(oFile);
		wcmbook = new XSSFWorkbook(input);

		wcmdataSheet = wcmbook.getSheet("WCM Content");	
		
		int rowNum = wcmdataSheet.getLastRowNum();
		String cellVal;
		CellStyle style = wcmbook.createCellStyle();// *
		Font font = wcmbook.createFont();// *

		for (int i = 0; i < rowList.size(); i++) {
			Map<String, String> map = new HashMap<String, String>();
			map = rowList.get(i);
			int cellNumber = 0;
			XSSFRow row = wcmdataSheet.createRow(++rowNum);
			for (Map.Entry<String, String> entry : map.entrySet()) {
				cellVal = entry.getValue();
				// System.out.println("Current Cell: " + cellNumber);
				XSSFCell cell = row.createCell(cellNumber);
				// System.out.println(cell + "-" + cellNumber);

				if (cellVal instanceof String) {
					cell.setCellValue((String) cellVal);
				} else {
					cell.setCellValue("String");
				}

				// dataSheet.autoSizeColumn(cellNumber);
				wcmdataSheet.setColumnWidth(0, 5000);
				wcmdataSheet.setColumnWidth(1, 13000);
				wcmdataSheet.setColumnWidth(2, 5000);
				wcmdataSheet.setColumnWidth(3, 13000);
				wcmdataSheet.setColumnWidth(4, 13000);
				wcmdataSheet.setColumnWidth(5, 5000);
				style.setWrapText(true);
				cell.setCellStyle(style);

				cellNumber++;
			}
		}
		input.close();
		FileOutputStream fos = new FileOutputStream(oFile);
		wcmbook.write(fos);
		wcmbook.close();
		fos.close();
	}

	
}
