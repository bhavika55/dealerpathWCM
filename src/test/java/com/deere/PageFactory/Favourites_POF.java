package com.deere.PageFactory;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Set;

import org.jsoup.Connection.Base;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.How;
import org.openqa.selenium.support.ui.Select;
import org.testng.Assert;
import org.testng.annotations.Test;
import org.testng.asserts.SoftAssert;

import com.deere.Helpers.BaseClass;
import com.deere.Helpers.GenericFactory;
import com.deere.Helpers.LogFactory;
import com.deere.Helpers.ReportFactory;
import com.deere.Helpers.ValidationFactory;

public class Favourites_POF {
	final WebDriver FavouritesDriver;
	static SoftAssert softAssert = new SoftAssert();

	public Favourites_POF(WebDriver driver) {

		this.FavouritesDriver = driver;
	}

	@FindBy(how = How.XPATH, using = ".//*[@id='favorites-target']/div/div")
	static WebElement wbelNofavaddedFav;

	@FindBy(how = How.XPATH, using = ".//*[@id='layoutContainers']/div/div[2]/div[3]/div[4]/section/div/div[2]/div[1]/h3")
	static WebElement wbelHeaderTitleFav;
	
	@FindBy(how = How.XPATH, using = ".//*[@id='layoutContainers']/div/div[2]/div[3]/div[4]/section/div/div[2]/div[1]/h3")
	static WebElement header_title_Favourites;

	@FindBy(how = How.ID, using = "Favourites-filter")
	static WebElement searchbox_Favourites;

	@FindBy(how = How.XPATH, using = ".//*[@id='Favourites-target']/div/div")
	static WebElement nofavadded_Favourites;

	@FindBy(how = How.XPATH, using = ".//*[@id='Favourites-target']/div/p[2]/span")
	static WebElement icon_Favourites;

	@FindBy(how = How.XPATH, using = ".//*[@id='Favourites-target']/div/p[2]")
	static WebElement message_Favourites;

	@FindBy(how = How.LINK_TEXT, using = "Actions")
	static WebElement action_Favourites;

	@FindBy(how = How.XPATH, using = "//input[@id='favorites-filter']")
	   static WebElement FavFilter;

	@FindBy(how = How.XPATH, using = ".//*[@class='section-body' and @id='Favourites-target']")
	static WebElement sectionBody;

	@FindBy(how = How.XPATH, using = ".//*[@id='leftNav']/li/a[@class='active']")
	public static List<WebElement> ListAllActiveLinks;	    

    @FindBy(how = How.XPATH, using = "//div[@class='section-header']//h3")
    public static List<WebElement> header;

    @FindBy(how = How.XPATH, using = "//*[@class='app-title']")
    static WebElement home;
    
	@FindBy(how = How.XPATH, using = ".//*[@id='js-favorites']")
	static WebElement wbelFavourSearch;
	  
	   /*
		 * @FindBy(how = How.LINK_TEXT, using = "Add Folder") static WebElement
		 * addfolder_Favourites;
		 */

		@FindBy(how = How.XPATH, using = ".//*[@id='Favourites-target']/div/div/div/div[1]/div/div[1]")
		static List<WebElement> favouritesFoldername;

		@FindBy(how = How.XPATH, using = "//a[@id='js-fav-actions-trigger']")
		static WebElement addfolderFavourites;

		@FindBy(how = How.XPATH, using = ".//*[@id='links-target']")
		static WebElement framePath;

		@FindBy(how = How.XPATH, using = ".//*[@class='icon fav-star is-selected']")
		static String selectedFavouriteicon;

		@FindBy(how = How.XPATH, using = ".//*[@id='links-target']/div/div/div/div/span")
		static List<WebElement> Favouriteicon;

		@FindBy(how = How.XPATH, using = "//div[@id='js-Favourites']")
		static WebElement favouriteQuickIcon;

		@FindBy(how = How.XPATH, using = "//div[@id='quick-Favourites-target']//div[contains(@class,'link-item')]")
		static List<WebElement> favouriteQuickLinkFolder;
		
		@FindBy(how = How.XPATH, using = ".//*[@class='icon fav-star']")
		static WebElement favouriteUnmarkedIconStarOnDepartmentpage; 
		
		@FindBy(how = How.XPATH, using = ".//*[@name='add-Favourite-folder']")
		static WebElement FavouriteAddFolderSelectboxpath;
		
		@FindBy(how = How.XPATH, using = ".//*[@id='add-Favourite-new-folder']")
		static WebElement FavouriteAddNewFolderSelectboxpath;
		
		@FindBy(how = How.XPATH, using = ".//*[@id='add-folder-modal']/div/div/div[3]/button[1]")
		static WebElement FavouriteAddFolderCancelbtn;
		
		@FindBy(how = How.XPATH, using = ".//*[@id='add-favorite-save']")
		static WebElement favouriteSavebtn;
		
		@FindBy(how = How.XPATH, using = ".//*[@class='icon fav-star is-selected']")
		static WebElement favouriteMarkedIconStarOnDepartmentpage;		
		
		@FindBy(how = How.XPATH, using = ".//*[@class='icon folder closed']")
		static List<WebElement> favouriteFolderClosedIcon;
		
		@FindBy(how = How.XPATH, using = "//div[@class='popover-content']//div[not(contains(@class,'link-item child  collapsed')) and contains(@class,'link-item')]")
		static List<WebElement> favouriteQuickIconFolderAndLinksList;
		private static Object createFavourite;
		
		@FindBy(how = How.XPATH, using = ".//*[@id='links-target']/div/div/div/div/span")
		static List<WebElement> favouiteIcon;
        @FindBy(how = How.XPATH, using = ".//*[@id='links-target']/div/div/div/div")
		static List<WebElement> favouiteLinkList;

		
		@FindBy(how = How.XPATH, using = ".//*[@id='js-fav-action-add-folder']")
		static WebElement addFolderbtn;
		
		@FindBy(how = How.XPATH, using = ".//*[@id='add-folder-name']")
		static WebElement addFolderNametxtbox;
		
		@FindBy(how = How.XPATH, using = ".//*[@id='add-folder-add']")
		static WebElement addbtn;		
		
		@FindBy(how = How.XPATH, using = ".//*[@id='add-folder-error-duplicate']")
		static WebElement addFolderDuplicate;		
		
		@FindBy(how = How.XPATH, using = ".//*[@id='leftNav']/li/a[@class='active']")
	
		public static List<WebElement> ActivedepartmentList;
		
		@FindBy(how = How.XPATH, using = ".//*[@id='js-fav-actions-trigger']")
		static WebElement favouiteActions;
		@FindBy(how = How.XPATH, using = "	.//*[@id='js-fav-action-add-folder']")
		static WebElement favouiteActions_JDIN;
	
		
		@FindBy(how = How.LINK_TEXT, using = ".//*[@class='custom-search-filter-list']/li[1]")
		static WebElement favActionsAddFolder;
		
		@FindBy(how = How.XPATH, using = "//span[@class='icon delete']")
		static List<WebElement> deleteIcon;
		
		@FindBy(how = How.XPATH, using = ".//*[@id='main-header']/div[1]/div/h1")
		static WebElement homepagepath;
		
		@FindBy(how = How.XPATH, using = ".//*[@id='layoutContainers']/div/div[2]/div[3]/div[4]/section/div/div[2]/div")
		static List<WebElement> favouriteTableElements;
		
		@FindBy(how = How.XPATH, using = ".//*[@id='favorites-target']/div/p[2]")
		static WebElement wbelMessageFav;

		@FindBy(how = How.XPATH, using = ".//*[@id='favorites-target']/div/div/div/div/div/div[1]")
		static List<WebElement> lstActualFoldernames;
		
		@FindBy(how = How.XPATH, using = "//div[@id='js-segments']")
		 static WebElement wbelProductSegmentIcon;
		
		@FindBy(how = How.XPATH, using = ".//*[@name='add-favorite-folder']")
 	    static WebElement favoriteAddFolderSelectboxpath;
		  
		@FindBy(how = How.XPATH, using = ".//*[@id='add-favorite-new-folder']")
	    static WebElement favoriteAddNewFolderSelectboxpath;
		
		@FindBy(how = How.XPATH, using = ".//*[@id='js-quick-favorites-popover']/div[2]/div[1]")
		static WebElement wbelQuickLinkHeader;
		
		
// *********************************************** XPATHS ********************************************************************************
	   
		
		
		// Return method to verify announcement table present or not
				public static boolean verifyFavouriteTableOnHomePage() throws Throwable {
					try {
						List<WebElement> favouriteTable_obj = favouriteTableElements;

						if (favouriteTable_obj.size() > 0) {

							return true;
						} else {
							return false;
						}
					} catch (Exception e) {
						// LogFactory.error("e");
						return false;
					}

				}
				
		
		
			   
				// Generic Method 
				public static ArrayList<String> ToGetFavouritesFolder_HomePage()
				{
					ArrayList<String> FavHomeFolderNameList = new ArrayList<String>();
					String FolderName="";
					
					try {
						// homepagepath.click();
						 List<WebElement> FolderListObj = BaseClass.wbDriver.findElements(By.xpath(".//*[@class='fav-link-body']"));
						
						 for(int i=0;i<FolderListObj.size();i++){
							 
							 WebElement folderobj = FolderListObj.get(i);
							 List<WebElement> linkobj = folderobj.findElements(By.tagName("a")) ;
							 if(!(linkobj.size()>0)){
								 
								 FolderName= folderobj.getText().toString().trim();
								 FavHomeFolderNameList.add(FolderName);
								 
							 }
						
						}
						LogFactory.info("Folder List On Home Page Portlet "+FavHomeFolderNameList);
						
					} catch (Exception e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
					return FavHomeFolderNameList;
				}
				// ***********************************END ***************************
				public static ArrayList<String> TOGetFolderListQuickModal()
				{
				
					ArrayList<String> QuickFolderNameList = new ArrayList<String>();
					String FolderName="";
					
					try {
						homepagepath.click();
						BaseClass.wbDriver.findElement(By.xpath(".//*[@id='js-favorites']")).click();
						List<WebElement> Quicklemetlist = BaseClass.wbDriver.findElements(By.xpath(".//*[@id='quick-favorites-target']/div/div/div/div"));
						
						 for(int i=0;i<Quicklemetlist.size();i++)
						 {
							 
							 WebElement folderobj = Quicklemetlist.get(i);
							 List<WebElement> linkobj = folderobj.findElements(By.tagName("a")) ;
							 if(!(linkobj.size()>0)){
								 
								 FolderName= folderobj.getText().toString().trim();
								 QuickFolderNameList.add(FolderName);
								 
							 }
						
						}
						 BaseClass.wbDriver.findElement(By.xpath(".//*[@id='js-favorites']")).click();
						LogFactory.info("Folder List under Quick Modal Window  "+QuickFolderNameList);
						
					} catch (Exception e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
					return QuickFolderNameList;
					
				}
				
				
			//*********************************** END ***************************************************************
				
				// Generic method to Get markedFav Link on Home Page My dealer Path

				public static ArrayList<String> ToGetMarkedFavouriteslink_HomePage()
				{
					ArrayList<String> DealerPathFavouriteslinkList = new ArrayList<String>();
					try {
						
						//BaseClass.wbDv.findElement(By.xpath(".//*[@id='js-Favourites']")).click();
						String FavlinkText="";
						//To read the Quick list
						// homepagepath.click();
						List<WebElement> DealerPathFavouriteslink = BaseClass.wbDriver.findElements(By.xpath(".//*[@class='fav-link-container']/div/a"));
						
						for(int i=0;i<DealerPathFavouriteslink.size();i++)
						{
							FavlinkText=DealerPathFavouriteslink.get(i).getAttribute("textContent").trim();
							DealerPathFavouriteslinkList.add(FavlinkText);	
						}
						LogFactory.info("Marked Favourite Link on Homepage Portlet"+DealerPathFavouriteslinkList);
						return DealerPathFavouriteslinkList;
					} catch (Exception e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
					return DealerPathFavouriteslinkList;
				}
					
			//*********************************** END ***************************************************************
					
				// Favourites should be change on preffered product type.
						// In this method we are verifying the Favoire link on changing the product type
								
						public static void ToSelectProductType(String ProductSegment)
						{
							// Click On Product Segmnets 
									
							 BaseClass.wbDriver.findElement(By.xpath(".//*[@id='js-segments']")).click();
							List<WebElement> Webproducttype = BaseClass.wbDriver.findElements(By.xpath(".//*[@class='checkbox-value']/div/label/input"));
									
							for(int i=0;i<Webproducttype.size();i++)
							{
								String productname = Webproducttype.get(i).getAttribute("id").toString().trim();
								// Unselecting the Product type which are not in product segmnet list
								if(ProductSegment.contains(productname))
								{
									Webproducttype.get(i).click();
								}
										
										
							}
								// Click On Apply Filter 
								 BaseClass.wbDriver.findElement(By.xpath(".//*[@id='js-segments-popover']/div[2]/div[3]/button")).click();
									
									
						}
						
						//************************************************************
						
		 public static void ToDeleteFavFolder(String folderName)
		{
			   ArrayList<String> actualfoldernameList = ToGetFavouritesFolder_HomePage();
			
		       
		      try {
		    	  //  to check folder is present 
		    	  if(actualfoldernameList.contains(folderName))
		    	  {
				// common Xpath to select Delete Button 
		    	   
		    	  
				    List<WebElement> deleteIconList = BaseClass.wbDriver.findElements(By.xpath("//span[@data-original-title='Delete Folder']"));
				          
				        for (int j = 0; j < actualfoldernameList.size(); j++)
				        {
				          String actualFolerName = actualfoldernameList.get(j).trim();
				           Thread.sleep(2000);
				            
				               if (actualFolerName.equals(folderName))
				                {            
				               @SuppressWarnings("unused")
				               
				               //to select the web element of delete btn
				               WebElement WebDeleteButtonobj = deleteIconList.get(j);          
				               
				               //move the cursor in corresponding foldername index
				               Actions Action_obj = new Actions(BaseClass.wbDriver);
				               Action_obj.moveToElement( WebDeleteButtonobj).build().perform();
				              Thread.sleep(3000);
				        
				               
				                JavascriptExecutor js = (JavascriptExecutor)BaseClass.wbDriver;
				                js.executeScript("arguments[0].click();", WebDeleteButtonobj);
				                
				                // Now click on Ok Button 
				                Thread.sleep(2000);
				               BaseClass.wbDriver.findElement(By.xpath(".//*[@id='delete-folder-ok']")).click();
				                 
				                  break;
				                // After delete Check Folder is Exist or not or home page
				                }
				               
				            }
				           
				        ArrayList<String> FolderList = ToGetFavouritesFolder_HomePage();
				           
				       // After deletion Folder list should not contaion the deleted Folder in List	               
				       if(!FolderList.contains(folderName))
				       {				    	 
				          LogFactory.info("Folder is deleted sucussfully "+folderName);
				          				          
				       }
				       else
				       {				    	 
				          LogFactory.info("Folder is not deleted sucussfully "+folderName);
				       }
		    	  }
		    	  else{
		    		  // folder is not present
		    	
		    		  LogFactory.info(" Given Folder is not not  present "+folderName);
		    	  }
		    	  // report 
		    	  
			} catch (InterruptedException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		            
		}
						
			//***************************  End Method *******************************************	
						
						// Generic method to Get markedFav Link on Dept Page

		public static ArrayList<String> ToGetMarkedFavouriteslink_Department()
			{
			
				ArrayList<String> selectedFavouriteIconLink = new ArrayList<String>();
				ArrayList<String> UnselectedFavouriteIconLink = new ArrayList<String>();
				String linktext="";
			try {
					List<WebElement> Favicon = BaseClass.wbDriver.findElements(By.xpath(".//*[@id='links-target']/div/div/div/div/span"));
				List<WebElement> Link = BaseClass.wbDriver.findElements(By.xpath(".//*[@id='links-target']/div/div/div/div"));
				
				
				for(int i= 0;i<Favicon.size();i++)
				{
					  String Attributevalue = Favicon.get(i).getAttribute("class").toString().trim();
				      linktext=Link.get(i).getText().toString().trim();
				         if(Attributevalue.contains("is-selected"))
					     {
						  selectedFavouriteIconLink.add(linktext);
					     }
					    else {
					    	UnselectedFavouriteIconLink.add(linktext);
					    }
					
					
				}
				LogFactory.info("Selected link with Favourite Star Icon on Department ---"+selectedFavouriteIconLink);
				//LogFactory.info("UnSelected link with Favourite Star Icon ---"+UnselectedFavouriteIconLink);
					
				
		} catch (Exception e) {
			// TODO Auto-generated catch block
				e.printStackTrace();
			}
			return selectedFavouriteIconLink;
		}

						//*********  End Method *******************************************
						
						
						// Generic method to Get UnmarkedFav Link on Dept Page
		public static ArrayList<String> ToGetUnMarkedFavouriteslink_Department()
		{

			ArrayList<String> selectedFavouriteIconLink = new ArrayList<String>();
			ArrayList<String> UnselectedFavouriteIconLink = new ArrayList<String>();
			String linktext="";
		try {
					List<WebElement> Favicon = BaseClass.wbDriver.findElements(By.xpath(".//*[@id='links-target']/div/div/div/div/span"));
				    List<WebElement> Link =  BaseClass.wbDriver.findElements(By.xpath(".//*[@id='links-target']/div/div/div/div"));
				
				    for(int i= 0;i<Favicon.size();i++)
				      {
					    String Attributevalue = Favicon.get(i).getAttribute("class").toString().trim();
				        linktext=Link.get(i).getText().toString().trim();
				             if(Attributevalue.contains("is-selected"))
					           {
						           selectedFavouriteIconLink.add(linktext);
					           }
					          else {
					    	         UnselectedFavouriteIconLink.add(linktext);
					                  }
							}
				//LogFactory.info("Selected link with Favourite Star Icon ---"+selectedFavouriteIconLink);
				LogFactory.info("UnSelected link with Favourite Star Icon On department ---"+UnselectedFavouriteIconLink);
					
			
		} catch (Exception e) {
			// TODO Auto-generated catch block
				e.printStackTrace();
			}
			return UnselectedFavouriteIconLink;
		}

		//****************************************
						
						// To Get Fav link on deprtmnet marked and unmarked both

		public static ArrayList<String> ToGetFavouriteslink_Department()
		{

			ArrayList<String> FavouriteLinkNameList = new ArrayList<String>();

			String linktext="";
		try {

				List<WebElement> LinkList =  BaseClass.wbDriver.findElements(By.xpath(".//*[@id='links-target']/div/div/div/div"));
			
			     for(int i= 0;i<LinkList.size();i++)
			      {
					linktext=LinkList.get(i).getAttribute("textContent").trim();
				    FavouriteLinkNameList.add(linktext);
				  }
		    	LogFactory.info("Fav Link on Departmnet is + ---"+FavouriteLinkNameList);

		} catch (Exception e) {
			// TODO Auto-generated catch block
				e.printStackTrace();
			}
			return FavouriteLinkNameList;
		}


		//*************************************************************
				//*********  End Method *******************************************
						
						
						// Generic Method to get Fav Link Name on marked My DealerPath Favourites
		public static ArrayList<String> ToGetMarkedFavouriteslinkQuickModal()
		{
			ArrayList<String> FavlinkNameList = new ArrayList<String>();
			String Linknmae="";
		  try {
		
			    BaseClass.wbDriver.findElement(By.xpath(".//*[@id='js-favorites']")).click();
		        List<WebElement> MarkedFAvlinkpath = BaseClass.wbDriver.findElements(By.xpath(".//*[@id='quick-favorites-target']/div/div/div/div/a"));
			          for(int i=0;i<MarkedFAvlinkpath.size();i++)
			          {
				        Linknmae= MarkedFAvlinkpath.get(i).getAttribute("textContent").trim();
				        FavlinkNameList.add(Linknmae);
			          }
			          BaseClass.wbDriver.findElement(By.xpath(".//*[@id='js-favorites']")).click();
		       LogFactory.info("Favourite Link on QuickModal Window "+FavlinkNameList);
			} catch (Exception e) 
		      {
			// TODO Auto-generated catch block
				e.printStackTrace();
			  }
		    BaseClass.wbDriver.findElement(By.xpath(".//*[@id='js-favorites']")).click();

			return FavlinkNameList;
		}


						//**********************************  End Method *******************************************
								
						
						
				
// *********************************************** GENERIC METHODS REQUIRED FOR FAVOURITE  ********************************************************************************	   
	
				
				
				
				
				
				
				
				
				
	
/*	*//**
	 * @Functionality To verify the presence of Favourites header on homepage.
	 * @throws Throwable
	 *//*
	public static void verifyFavouritesHeaderPresent(String strexpectedheadermsg) throws Throwable {
		
		String TCID = "TC01_Favourite";
		String flag = "Fail";
		String result = "Favourites portlet with header text is not displayed.";
		
		try {
			String stractualheadermsg = Favourites_POF.header_title_Favourites.getText().toString();
			// System.out.println(stractualheadermsg);

			result = "Favourites portlet with header text is not displayed.";

			if (ValidationFactory.isElementPresent(header_title_Favourites)) {

				if (stractualheadermsg.equals(strexpectedheadermsg)) {

					flag = "Pass";
					result = "Favourites portlet with header text is displayed.";
					LogFactory.info("Favourites portlet with header is displayed.");

				}
			}
			ReportFactory.ReporterOutput(TCID, "Verify Favourites portlet with header text is displayed on homepage.",
					strexpectedheadermsg, "Favourites portlet header should be " + strexpectedheadermsg, result, flag);

			if (flag.equalsIgnoreCase("FAIL")) {
				Assert.assertFalse(true);
			}
		}

		catch (Exception e) {
			
			// LogFactory.error("e");
			String er = e.getMessage().substring(0, 25).toString().trim();

			ReportFactory.ReporterOutput(TCID, "Favourites portlet with header on homepage ", "NA", "NA", er, flag);
		}
	}

	*//**
	 * @Functionality To verify No Favourites added yet message on homepage.
	 * @throws Throwable exception.
	 *//*
	public static void verifyWhenNoFavouritePresent(String strexpectednofavmsg) throws Throwable {
		
		String TCID = "TC01_Favourite";
		String flag = "Fail";
		String result = "Favourites portlet is not displayed ";
		
		try {
			String stractualnofavmsg = Favourites_POF.nofavadded_Favourites.getText().toString();
			// System.out.println(stractualnofavmsg);

			flag = "Pass";
			result = "Message 'No Favourites added yet' is displayed ";
			
			if (ValidationFactory.isElementPresent(nofavadded_Favourites)) {

				if (stractualnofavmsg.equals(strexpectednofavmsg)) {
					flag = "Pass";
					result = "Message 'No Favourites added yet' is displayed.";
					LogFactory.info("Verify 'No Favourites added yet.' message is displayed.");

				}
			}

			ReportFactory.ReporterOutput(TCID,"Verify 'No Favourites added yet' message on home page when No Favourites added.",
					strexpectednofavmsg, "Message 'No Favourites added yet' should be displayed " + strexpectednofavmsg,
					result, flag);

			

		} catch (Exception e) {
			ReportFactory.ReporterOutput("Error_No_Favourite_added.",
					"Verify 'No Favourites added yet' message on home page when No Favourites added.",
					"Message 'No Favourites added yet' should be displayed ", "NA", e.getMessage().toString(), flag);
		}
	}

	
	*//**
	 * @param departmentName 
	 * @param ContentType 
	 * @param TCID 
	 * @param TCID 
	 * @return 
	 * @Functionality : To verify favourite icon with message on homepage.
	 * @throws Throwable
	 *//*
	public static void verifyFavouriteIconWhenNoFavouriteAdded(String strexpectedicon) throws Throwable {
		
		String TCID = "TC01_Favourite";
		String flag = "Fail";
		try {
			String result = "Star icon with Message is not displayed.";

			if (ValidationFactory.isElementPresent(message_Favourites)) {
				String stractualicon = Favourites_POF.message_Favourites.getText().toString();
				// System.out.println(stractualicon);

				if (stractualicon.equals(strexpectedicon)) {
					LogFactory.info("Verify Star icon with Message is displayed.");
					flag = "Pass";
					result = "Star icon with Message is displayed.";
				}
			}

			ReportFactory.ReporterOutput(TCID,
					"Verify Favourites Icon with message in Favourites portlet when no favourites", strexpectedicon,
					"Star icon with Message should be  " + strexpectedicon, result, flag);

			if (flag.equalsIgnoreCase("FAIL")) {
				Assert.assertFalse(true);
			}

		} catch (Exception e) {

			ReportFactory.ReporterOutput("Error_Favourites", "Favourites Icon with Message.",
					"Favourites Icon with Message", "NA", e.getMessage().toString(), flag);
		}

	}

*/
				
				
//*********************************************** Completed by Archana *****************************************************************************************************************************

// **************************************************** E - N - D **********************************************************************************************************

			
				
	//[P26-- Verify creating Favourites on homepage.]
		public static ArrayList<String> createFavouriteOnDepartment(String TCID, String departmentName) throws Throwable
	       {
	             String flag = "Fail";
	             String result= "Favourite is not created ";
	                                 
	             String actualUtilityLinkLists = "";
	             String expectedDataValue = "";
	             String linknmae="";       
	             int selectedindex=0;
	             String IconIsslected ="";
	             
	  ArrayList<String> markedFavlink = new ArrayList<String>();
	             
	  try{                
	      
	          //To clik on department page                          
	         if( GenericFactory.toClickonDeptOnFavourite(departmentName))
	         {
	          Thread.sleep(2000);
	                                                                  
	          //To clik on star icon of an link to create favourite
	         List<WebElement> favouriteIconList = favouiteIcon;
	         List<WebElement> favouriteLinkList = favouiteLinkList;
	                                                        
	         ArrayList<String> Favlinklist = ToGetFavouriteslink_Department();
	         if(Favlinklist.size()>0)
	         {
	                                              
	               int count = 5;
	               int maxlimit=0;
	              int Favlinklistcount = Favlinklist.size();
	              if(Favlinklistcount>count)
	              {
	                   maxlimit=count;
	              }
	              else{
	                   maxlimit=Favlinklistcount;
	              }
	                                                                        
	        //&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&                          
	        for(int i=0;i<maxlimit;i++)
	        {    
	        	selectedindex=0;
	            IconIsslected = "";
	           linknmae=favouriteLinkList.get(i).getAttribute("textContent").trim();
	            IconIsslected = favouriteIconList.get(i).getAttribute("class").toString().trim();
	           //if(linknmae.contains(FavLink.trim()))
	            markedFavlink.add(linknmae);
	                                
	         //{                                                        
	          if(IconIsslected.contains("fav-star is-selected"))
	           {
	        	  LogFactory.info(linknmae +" is Already Marked so First Unmark the link and Create Marked fav");
	        	  favouriteIconList.get(i).click();
	        	  Thread.sleep(1000);
	           }
	        	
	            
	             favouriteIconList.get(i).click();
	              selectedindex=i;
	             System.out.println("clicked on the link to mark- " +linknmae);
	             Thread.sleep(6000);                             
	           
	         //click on save btn
	            if(favouriteSavebtn.isDisplayed()){
	            	  favouriteSavebtn.click();
	            }
	       
	         Thread.sleep(6000);
	              
	         // Verify after  clicked on Fav , its selected or not.
	         IconIsslected = favouriteIconList.get(selectedindex).getAttribute("class").toString().trim();
	         String txt11 = favouriteLinkList.get(selectedindex).getAttribute("textContent").trim();
	         if(IconIsslected.contains("fav-star is-selected"))
	         { 
	               flag = "Pass";
	              result= "Favourite is created ";                                                                           
	            //Pass report
	            LogFactory.info("Link is marked as favourite "+linknmae);
	                                                                                                  
	         }
	         else
	         {
	              flag = "Fail";
	             result= "Unable to marked the link as Favourite "+linknmae;                                                                             
	           //Fail report
	            LogFactory.info("Unable to marked the link as Favourite "+linknmae);
	            ReportFactory.reporterOutput(TCID, "Verify creating a favorite.", 
  		              "Created fav Links= "+markedFavlink,"Favorite should be display in yellow color on Quick modal window on home, department and outside a Favorite folder. ",
  		                                              result, flag);
	            break;
	         }            
	     
	         
	                                                                                                                                         
	       }
	         }                                                       
	        else{
	             flag = "Pass";
	             result= "Favouritelink is not present so can not marked links as a favourite ";
	             ReportFactory.reporterOutput(TCID, "Verify creating a favorite.", 
   		              "Created fav Links= "+markedFavlink,"Favorite should be display in yellow color on Quick modal window on home, department and outside a Favorite folder. ",
   		                                              result, flag);
	           }
	         }
	         else{
	        	 flag = "Pass";
	             result= "Departmnent is not clickable so can not create the Link as favourite "+departmentName;
	             ReportFactory.reporterOutput(TCID, "Verify creating a favorite.", 
	            		 departmentName+ " is not clickable","Favorite should be display in yellow color on Quick modal window on home, department and outside a Favorite folder. ",
   		                                              result, flag);
	         }
	     
	   // Verification of created Favorite on Home page ,Quick Favorite modal window on home, department        
	                         
	    
	        if(flag=="Pass" &&markedFavlink.size()>0)
	         {
	        	
	        	    ArrayList<String> QuickFavlink= new ArrayList<String>();
	        	    ArrayList<String> FavlinkHompegae= new ArrayList<String>();
	        	    
		        	// Verify Created link is present on homepage
	        	      homepagepath.click();
		              FavlinkHompegae = ToGetMarkedFavouriteslink_HomePage();
		             
		        	   QuickFavlink = ToGetMarkedFavouriteslinkQuickModal();
		        	
		        	   if(FavlinkHompegae.containsAll(markedFavlink)&&QuickFavlink.containsAll(markedFavlink))
		        	   {
		        		     flag = "Pass";
		        		   result="Favorite  get created in the default Favorite folder and displayed on 'Favorites' portlet on homepage independently outside a Favorite folder and On Quick Model.";
		        		   LogFactory.info(result);
		        		   ReportFactory.reporterOutput(TCID, "Verify creating a favorite.", 
		        		              "Created fav Links= "+markedFavlink,"Favorite should be display in yellow color on Quick modal window on home, department and outside a Favorite folder. ",
		        		                                              result, flag);
		        		}
		        	   else{
		        		   result="Favorite  get created in the default Favorite folder and but not displayed on 'Favorites' portlet on homepage independently outside a Favorite folder and On Quick Model.";
		        		     flag = "Fail";
		        		     LogFactory.info(result);
		        		     ReportFactory.reporterOutput(TCID, "Verify creating a favorite.", 
		        		              "Created fav Links= "+markedFavlink,"Favorite should be display in yellow color on Quick modal window on home, department and outside a Favorite folder. ",
		        		                                              result, flag);
		        	   }
		        	   
		        		
		        	     
		        	     
	         }
	     
	                                                           
	     }catch (Exception e) 
	      {                  
	            String er = e.getMessage().substring(0, 60).toString().trim();
	             
	          ReportFactory.reporterOutput(TCID, "Verify creating a favorite.", "NA",
	        		  	"NA", er, flag);
	      }                              
	       // back to homepage
	      homepagepath.click();
	      return markedFavlink;
    }            

// **************************************************** E - N - D **********************************************************************************************************	
	
	//[P26-- Verify creating (Add) folder for Favourites from Favourite portlet on homepage.] 
	//[ Naveen helped]
	
	public static void verifyCreateFolder(String TCID, String folderName) throws Throwable
	{
		String flag = "Fail";
		String result= "Favourite Folder is not created ";
				
		try{		
				
				//Click On Add Folder Button
				homepagepath.click();
				if(ToGetFavouritesFolder_HomePage().contains(folderName))
				{
					 LogFactory.info("Favourite folder is already Exist so delete the folder First "+folderName);
					 ToDeleteFavFolder(folderName);
				}
					
				Thread.sleep(2000);
				if(favouiteActions.isDisplayed())
				{
					favouiteActions.click();
				}
				else
				{
					favouiteActions_JDIN.click();
				}
				Thread.sleep(2000);
	//[Naveen]	// Execute the Java Script for the element which we find out
				((JavascriptExecutor) BaseClass.wbDriver).executeScript("favActionsAddFolder();");
				
				//Enter Folder Name
				addFolderNametxtbox.sendKeys(folderName);
				addbtn.click();
				 Thread.sleep(2000);
				 
				// To check Folder is already exist or not
				 String errorobjtext = addFolderDuplicate.getText().toString().trim();			
				
				     if(!(errorobjtext.length()>0))
				     {				 		
					   // If does not Exist click on Save button 
					  
						// To Get Folder Name List after adding new folder
				    	 Thread.sleep(4000);
					     ArrayList<String> FolderNameListHomepage = ToGetFavouritesFolder_HomePage();
					     ArrayList<String> FolderNameListQuickModal = TOGetFolderListQuickModal();
					     
					     
						     if(FolderNameListHomepage.contains(folderName)&&FolderNameListQuickModal.contains(folderName))
						     {   	 // Pass
						    	 flag = "Pass";
						    	 result= "   Favourite folder created and displaying on Homepage, Quickmodal Homepage, Department  ";
						    	 LogFactory.info(result);
						     }
						     else
						     {  // fail
						    	 flag = "Fail";
						 		 result= "Favourite folder is not created ";
						 		 LogFactory.info(result);
						     }			   					  
				     }
				     else
				     {
				    	 // Folder already Exist
				    	 flag = "Fail";
				    	 result= "Favourite folder is already Exist ";	
				    	 LogFactory.info(errorobjtext);
				    	 
				    	 // Click on Cancel Button 
				    	 FavouriteAddFolderCancelbtn.click();
				     }
			
			
			
			ReportFactory.reporterOutput(TCID,"Verify creating folder for favorites from favorite portlet on homepage.", folderName,
					"   Favourite folder should get created and display on Homepage, Quickmodal, Department ' Portlet on homepage and on quick favourite modal window on home, department", result, flag);
		
		}catch(Exception e)
		{			
			String er = e.getMessage().substring(0, 160).toString().trim();

			ReportFactory.reporterOutput(TCID, "Verify creating a Favourite folder ", "NA",
					"NA", er, flag);    	
		}
		
		//To click on dealerpath homepge
		homepagepath.click();
	}
	
// **************************************************** E - N - D **********************************************************************************************************
	
	//Verify Deleting Folder in Favourite portlet on homepage.
	public static void verifyDeleteFavouriteFolder(String TCID, String folderName) throws Throwable
	{  
		String flag = "Fail";
		String result= "Favourite Folder is not deleted ";
		
		try{			
			
		   // To Get Folder Name List from Home Page
	     ArrayList<String> actualfoldername1 = ToGetFavouritesFolder_HomePage();
	     if(actualfoldername1.contains(folderName))
	        
	        {
	              
	       // common Xpath to select Delete Button 
	        List<WebElement> deleteIconList = deleteIcon;
	              
	            for (int j = 0; j < actualfoldername1.size(); j++)
	            {
	              String actualFolerName = actualfoldername1.get(j).trim();
	               Thread.sleep(2000);
		            
		               if (actualFolerName.equals(folderName))
		                {            
		               @SuppressWarnings("unused")
		               
		               //to select the web element of delete btn
		               WebElement WebDeleteButtonobj = deleteIconList.get(j);          
		               
		               //move the cursor in corresponding foldername index
		               Actions Action_obj = new Actions(BaseClass.wbDriver);
		               Action_obj.moveToElement( WebDeleteButtonobj).build().perform();
		              Thread.sleep(3000);
		        
		             //use javaScriptExceuter to control delete functioality
		                JavascriptExecutor js = (JavascriptExecutor)BaseClass.wbDriver;
		                js.executeScript("arguments[0].click();", WebDeleteButtonobj);
		                
		                // Now click on Ok Button 
		                Thread.sleep(2000);
		               BaseClass.wbDriver.findElement(By.xpath(".//*[@id='delete-folder-ok']")).click();
		                 LogFactory.info("Favourite folder is deleted. ");
		                  break;
		               
		                }
		               
	                }
	            // After delete Check Folder is Exist or not or home page
	            ArrayList<String> FolderListHomePage = ToGetFavouritesFolder_HomePage();
	            ArrayList<String> FolderListQuickModal= TOGetFolderListQuickModal();
	            
	               
               // After deletion Folder list should not contaion the deleted Folder in List	               
               if((!FolderListHomePage.contains(folderName))&&(!FolderListQuickModal.contains(folderName)))
               {
                  LogFactory.info("Folder is deleted sucussfully "+folderName);
                  flag= "Pass";
		    	  result= "Favourite folder removed and not displayed more on homepage and quick model";
                  //Pass report
                  
               }
               else{
                  // Fail Report
            	   flag= "Fail";
 		    	  result= "Folder is not deleted ";
                  LogFactory.info("Folder is not deleted sucussfully "+folderName);
               }
	        }
	     else{
	  	       flag= "Pass";
	    	  result= "Folder is not Present so can not be deleteed- "+folderName;
           LogFactory.info("Folder is not Present so can not be deleted- "+folderName);
	     }
               
	  ReportFactory.reporterOutput(TCID,"Verify deleting Favourite folder ", folderName,
				"The favorite deleted should not be available in Quick Favorite modal window, Favorite Portal on homepage ", result, flag);
	           
	    
	  
		}catch(Exception e)
		{			
			String er = e.getMessage().substring(0, 160).toString().trim();

			ReportFactory.reporterOutput(TCID, "Verify deleting Favourite folder  ", "NA",
					"NA", er, flag);    	
		}
	        }      

	
// **************************************************** E - N - D **********************************************************************************************************
	
	
	//"Verify removing a Favourite.
	public static void verifyRemoveFavouriteLink(String TCID, String departmentName, String Favlink) throws Throwable
	{	
		String flag = "Fail";
		String result= "Favourite Link is not removing ";
		String Isslected = "";
		String linknmae = "";
		
	try{			
		
			//To clik on department page
		if(GenericFactory.toClickonDeptOnFavourite(departmentName))
		{
						
				
			//To clik on star icon of an link to create favourite
			    ArrayList<String> MarkedFavlink = ToGetMarkedFavouriteslink_Department();
				List<WebElement> Favicon = BaseClass.wbDriver.findElements(By.xpath(".//*[@id='links-target']/div/div/div/div/span"));
				List<WebElement> Link = BaseClass.wbDriver.findElements(By.xpath(".//*[@id='links-target']/div/div/div/div"));
				int selectedindex=0;
			if(MarkedFavlink.contains(Favlink)&&Favlink.length()>0)
			{
				for(int i=0;i<Link.size();i++)
				{
					
					 linknmae = Link.get(i).getAttribute("textContent").trim();
					if(linknmae.contains(Favlink.trim())){
						 Isslected = Favicon.get(i).getAttribute("class").trim();
						
						if(Isslected.contains("fav-star is-selected")){
						     Favicon.get(i).click();
						     LogFactory.info("Clicked on link to remove "+Favlink);
						      selectedindex=i;
						     break;
						}					
					}				
				}
				//verify quick modal winow
				ArrayList<String> favlist = ToGetMarkedFavouriteslink_Department();
				homepagepath.click();
				Thread.sleep(2000);
				ArrayList<String> quickLinkNameList = ToGetMarkedFavouriteslinkQuickModal();
				
			     //Go to back hhomepage
				//clik on homepage
				homepagepath.click();
				ArrayList<String> favouiteLinkListOnHomepage = ToGetMarkedFavouriteslink_HomePage();
				
				if((!favouiteLinkListOnHomepage.contains(Favlink)) && (!quickLinkNameList.contains(Favlink))&&(!favlist.contains(Favlink)))
				{
					//pass
					 flag = "Pass";
					 result= "Favourite is removed and no more displaying on 'Favourite' Portlet on homepage,quick favourite modal window Home page and Department";
					 LogFactory.info("Favourite Link is removed ");
					ReportFactory.reporterOutput(TCID, "verify Remove Favourite Link from Link portlet Department ", Favlink,
							"Favourite should get removed and should no more display on 'Favourite' Portlet on homepage,quick favourite modal window page", 
							result, flag); 
	
				}
				else
				{
					//fail
					 flag = "Fail";
					 result= "Favourite is not  removed and  displaying on 'Favourite' Portlet on homepage,quick favourite modal window Home page and Department";
					LogFactory.info("Favourite Link is not removed ");
					ReportFactory.reporterOutput(TCID, "verify Remove Favourite Link from Link portlet Department ", Favlink,
							"Favourite should get removed and should no more display on 'Favourite' Portlet on homepage,quick favourite modal window page", 
							result, flag); 
				}				
		    }
			else{
				// Link is not Present So can noy
				    flag = "Pass";
				    result= Favlink+" Link is not present on Page so can not remove.";
				 
				    LogFactory.info(result);
					ReportFactory.reporterOutput(TCID, "verify Removing a Favourite Link ", Favlink,
							"The favorite removed should not be available in Quick Favorite modal window, Favorite Portal on homepage, and the particular link in Links Portlet should not have a yellow star. ", result, flag); 
				 
			}
		}
		
		 else
		 {
			    flag = "Pass";
			    result= "Departmnent is not clickable so Can not remove the link ";
			 
			    LogFactory.info(result);
				ReportFactory.reporterOutput(TCID, "verify Removing a Favourite Link ", Favlink,
						"The favorite removed should not be available in Quick Favorite modal window, Favorite Portal on homepage, and the particular link in Links Portlet should not have a yellow star. ", result, flag); 
			
		   }
		
		
			
			
	
		
	
		
	}catch(Exception e)
		{			
			String er = e.getMessage().substring(0, 60).toString().trim();

			ReportFactory.reporterOutput(TCID, "Verify removing a favourite link ", "NA",
					"NA", er, flag);    	
		}
		
		//To click on dealerpath homepge
		homepagepath.click();
				
			
		
	}
	
// **************************************************** E - N - D **********************************************************************************************************
	
	//Verify moving a Favourite link and Folder from Favourite portlet on homepage.
	public static void verifyMovingFavouriteFolderandLink(String TCID) throws Throwable
	{
		String flag = "Fail";
		String result= "Favourite folder and link is not moving up and Down ";
		    JavascriptExecutor js = (JavascriptExecutor) BaseClass.wbDriver;
		    js.executeScript("window.scrollBy(0,500)");
		
		try{
			if((ToGetFavouritesFolder_HomePage().size()>0 ) || ToGetMarkedFavouriteslink_HomePage().size()>0)
			{
			Actions action = new Actions(BaseClass.wbDriver);
			
			WebElement from = BaseClass.wbDriver.findElement(By.xpath(".//*[@class='icon drag ui-sortable-handle']"));
			WebElement to = BaseClass.wbDriver.findElement(By.xpath(".//*[@class='icon drag ui-sortable-handle']"));
			
			action.dragAndDrop(from, to).perform();
			System.out.println("Draging the Webelement from its destination to new destination ");
			LogFactory.info("Draging the Webelement from its destination to new destination ");
			
			 flag = "Pass";
			 result= "Favourite folder and link is moving up and Down ";
			}
			else
			{
				flag = "Pass";
				 result= "Favourite folder or link is not present so can not validate  ";
			}
			ReportFactory.reporterOutput(TCID, "Verify moving a favorite link and Folder from Favorite portlet on homepage ", "NA",
	 		       "Favourite and Favourite folder should be moved accordingly ", result, flag);
			
		}catch(Exception e)
		{
			String er = e.getMessage().substring(0, 100).toString().trim();
			
			ReportFactory.reporterOutput(TCID, "Verify moving a favorite link and Folder from Favorite portlet on homepage ", "NA",
		 		       "Favourite and Favourite folder should be moved accordingly ", er, flag);
		}
		
		//To click on dealerpath homepge
				homepagepath.click();
	}
	

	
// **************************************************** E - N - D **********************************************************************************************************
	
	
	//[ P29-- Verify Collapsing and Expanding a folder on Favourites portlet. ]
	
		public static void verifyExpandAndCollapseFavroiteFolder(String TCID) throws Throwable
		{
			String flag = "Fail";
			String result= "Expand and Collapse Favourite Folder is not working ";
			String FolderName="";
			
			try{
				   JavascriptExecutor js = (JavascriptExecutor) BaseClass.wbDriver;
				    js.executeScript("window.scrollBy(0,500)");
						
					if(ToGetFavouritesFolder_HomePage().size()>0)
					{
					
					  int ClosedfoldercountBeforeExpand = favouriteFolderClosedIcon.size();
					 
					
					  List<WebElement> FolderListObj = BaseClass.wbDriver.findElements(By.xpath(".//*[@class='fav-link-body']"));
						List<WebElement> Faviconlist = BaseClass.wbDriver.findElements(By.xpath(".//*[@class='fav-link-body']/span"));
						 for(int i=0;i<FolderListObj.size();i++)
						 {
							 
							 WebElement folderobj = FolderListObj.get(i);
							 List<WebElement> linkobj = folderobj.findElements(By.tagName("a")) ;
							 if(!(linkobj.size()>0))
							  {
								 
								 FolderName= folderobj.getText().toString().trim();
								 Faviconlist.get(i).click();
								 LogFactory.info("Clicked to Expand on folder-  "+FolderName);
								 
							  }	 
						   }
						 
						  int ClosedfoldercountAfterExpand = favouriteFolderClosedIcon.size();
						  if(ClosedfoldercountAfterExpand==0)
						     {
							  
							  LogFactory.info("All Folder are Expanded ");
							  
							  }
						  
						  // No End Impersonate and 
						  GenericFactory.endImpersonateOrLogoutUser();
						  GenericFactory.utilityMenuAdminClick();
						  GenericFactory.impersonateUser(BaseClass.strUserRACFID);
						  int ClosedfoldercountafterRelogin = favouriteFolderClosedIcon.size();
						  if(ClosedfoldercountafterRelogin==ClosedfoldercountBeforeExpand)
						  {
							 flag = "Pass";
							 result= "Expand and Collapse Favourite Folder is  working ";

						  }
						  else
						  {
							    flag = "Fail";
								 result= "Expand and Collapse Favourite Folder is not working "; 
						  }
						  
						  
					}
					else
					{
						flag = "Pass";
						result= "Favourite Folder is not present so Expand and Collapse can not be verified ";
					}
				
				
				
				
				
				ReportFactory.reporterOutput(TCID, "Verify Collapsing and Expanding a folder on Favorites portlet ", "NA",
		    		       "Favourite folder should be expanded or collapsed accordingly ", result, flag);
			
			}catch(Exception e)
			 {			
				String er = e.getMessage().substring(0, 160).toString().trim();

				ReportFactory.reporterOutput(TCID, "Verify Collapsing and Expanding a folder on Favorites portlet ", "NA",
						"NA", er, flag);		    	
			 }
			
			//To click on dealerpath homepge
			homepagepath.click();
		}
		
		
		
// **************************************************** E - N - D **********************************************************************************************************
		
		
	//Verify Quick Favourite modal window on Home, Department and all levels of Index pages.
	       public static void verifyQuickModalWindowOnHomepage(String TCID) throws Throwable
           {
                  String flag = "Fail";
                  String result= "Quick Link List and favourite Link List both are not same ";
                  ArrayList<String> favouriteList= new ArrayList<String>();
                  ArrayList<String> quickLinkList= new ArrayList<String>();
                  ArrayList<String> favouriteListOfDepartment= new ArrayList<String>();
                  
           try{
                  // To Get Fav List from Homepage -1 
                   favouriteList = ToGetMarkedFavouriteslink_HomePage();  
     
                  
                  //to Get Fav List from Quick Search- 2
                   quickLinkList = ToGetMarkedFavouriteslinkQuickModal();
                  flag="Fail";
                  result="Quick Link List and favourite Link List both are not same ";
                  
                  //compare both should be same 
                  
                  if(favouriteList.equals(quickLinkList))
                  {
                        flag="Pass";
                        result="Quick Link List and favourite Link List both are same ";
                        LogFactory.info("All links are present in 'My DealerPath Favorites' is same in quick modal window ");
                        ReportFactory.reporterOutput(TCID, "Verify quick favouite marked list", "NA", "My DealerPath Favorites' should be same as in quick modal window  ", result, flag);
                        
                        
                        //report O/P
                  }                   
                  else{
                        flag="Fail";
                        result="QuickLinkList and FavouriteList both are not same ";
                        LogFactory.info("All links are present in 'My DealerPath Favorites' is same as in quick modal window ");
                        
                        ReportFactory.reporterOutput(TCID, "Verify quick favouite marked list", "NA",
                        		"Actual Link list are= "+favouriteList,  "Quick modal list are= "+ " " +quickLinkList +" "+result, flag);
                        
                        
                        //report O/P
                  }
                  
                        
                  
                        // Count the depatmnet list
                  // go on each departmnet and get fav link and fav should be present in 1 and 2
                  
                  //To get all department 
                  List<WebElement> departmentListCount = ActivedepartmentList;
                  for(int i=1; i<departmentListCount.size(); i++)
                  {
                        //to click on department one by one
                        String departmentName = departmentListCount.get(i).getText().toString().trim();
                        
                        
                         if(GenericFactory.toClickonDeptOnFavourite(departmentName))
                        {
                               
                               favouriteListOfDepartment = ToGetMarkedFavouriteslink_Department();
                                     
                                      if((favouriteList.containsAll(favouriteListOfDepartment)&&favouriteListOfDepartment.size()>0))
                                     {
                                            flag= "Pass";
                                            result= "MarkedFavouriteslink_Department is also present on home page";
                                            LogFactory.info("Favorite link on "+departmentName + favouriteListOfDepartment +" is is present in home page link  "+favouriteList);
                                     }
                                     else if(!(favouriteListOfDepartment.size()>0))
                                     {
                                            LogFactory.info(" No Favorite link on  this "+ departmentName );
                                     }
                                     else{
                                            
                                            flag= "Fail";
                                            result= "MarkedFavouriteslink_Department is not present on home page";
                                            LogFactory.info("Favorite link on "+ departmentName + favouriteListOfDepartment +" is is present in home page link  "+favouriteList);
                                            break;
                                                                                         
                                     }                         
                        }                   
                  }
                  
           }catch(Exception e)
                  {                  
                        String er = e.getMessage().substring(0, 160).toString().trim();

                        ReportFactory.reporterOutput(TCID, "Verify marked link on departmenet and on home page ", "NA",
                                     "NA", er, flag);                     
                   }
                  
                  //To click on dealerpath homepge
                  homepagepath.click();
                                            
     }      






		
		//*******************************************END************************************
	       
	       public static void VerifyFilter( String TCID,String StrInputvalue) throws Throwable
           {   String FavlinkListHomepage= "";
               String FavFolderListHomepage= "";
               String flag = "Fail";
                    String result= "Favourite Filter is not working properly on home page on entering data ";                

              
             try {
                                FavlinkListHomepage = ToGetMarkedFavouriteslink_HomePage().toString();
                                 FavFolderListHomepage = ToGetFavouritesFolder_HomePage().toString();
                                if(FavlinkListHomepage.contains(StrInputvalue)||FavFolderListHomepage.contains(StrInputvalue))
                                {
                                      LogFactory.info("Verify Filter with valid input --"+StrInputvalue);
                                         Thread.sleep(1000);
                                         FavFilter.click();
                                        FavFilter.sendKeys(StrInputvalue);
                                       FavlinkListHomepage = ToGetMarkedFavouriteslink_HomePage().toString();
                                       FavFolderListHomepage = ToGetFavouritesFolder_HomePage().toString();
                                          if(FavlinkListHomepage.contains(StrInputvalue)||FavFolderListHomepage.contains(StrInputvalue))
                                          {
                                          // Pass with Valid Data
                                          flag = "Pass";
                                            result= "Favourite Filter is working properly on home page on entering Valid data "; 
                                            LogFactory.info(result);
                                          }
                                          else{
                                          // Fail with Valid Data
                                            flag = "Fail";
                                                  result= "Favourite Filter is not working properly on home page on entering Valid data ";
                                                  LogFactory.info(result);
                                          }
                                      
                                }
                              else
                              {     LogFactory.info("Verify Filter with invalid input --"+StrInputvalue);
			                              Thread.sleep(1000);
			                              FavFilter.click();
                                       FavFilter.sendKeys(StrInputvalue);
                                       Thread.sleep(1000);
                                     FavlinkListHomepage = ToGetMarkedFavouriteslink_HomePage().toString();
                                     FavFolderListHomepage = ToGetFavouritesFolder_HomePage().toString();
                                     if(FavlinkListHomepage.contains(StrInputvalue)||FavFolderListHomepage.contains(StrInputvalue))
                                     {
                                            flag = "Fail";
                                            result= "Favourite Filter is not working properly on home page on entering InValid data ";
                                            LogFactory.info(result);
                                      }
                                      else
                                      {
                                            flag = "Pass";
                                            result= "Favourite Filter is working properly on home page on entering InValid data ";
                                            LogFactory.info(result);
                                      }
                                      
                               }
                               
                                  ReportFactory.reporterOutput(TCID, "Verify filtering of the favorite portlet ",
                                           StrInputvalue, "No favorites should be displayed (filtered) with no error messages for Invalid Data and Result should display for Valid Data", result, flag);
                       } catch (Exception e) {
                              String er = e.getMessage().substring(0, 125).toString().trim();

                              ReportFactory.reporterOutput(TCID, "Verify filtering of the favorite portlet  ", "NA",
                                           "NA", er, flag);  
                       }
             homepagepath.click();
           }
            


// **************************************************** E - N - D **********************************************************************************************************	
					
					
	//Verify Favourites for different languages.
				
		//Verify Favourites for different languages on homepage.
		public static void verifyFavouritesDifferentLanguagesOnHomepage(String TCID, String expectedFavouriteLink) throws Throwable
		{	
			String flag = "Fail";
			String result= "Favourites are not showing in user preferred languages on homepage ";
				
			try{
					//String expectedFavouriteLink = "GSA Emp OEM Link Test";
				ArrayList<String> FavlinkHomepage = ToGetMarkedFavouriteslink_HomePage();
				ArrayList<String> FavlinkQuickModal= ToGetMarkedFavouriteslinkQuickModal();
				
				if(FavlinkHomepage.size()>0 &&expectedFavouriteLink.length()>0)
				{
					if(FavlinkHomepage.contains(expectedFavouriteLink)&&FavlinkQuickModal.contains(expectedFavouriteLink))
					{
						flag = "Pass";
						result= "All the favorites is displayed for the site irrespective of the languages selected on Favorites portlet on homepage and quick favorite modal window on home";						
						LogFactory.info("Favourite are in on home page "+FavlinkHomepage +" " +expectedFavouriteLink );
						
						
					}
					else
					{
						flag = "Fail";
						result= "All the favorites is not  displayed for the site irrespective of the languages selected on Favorites portlet on homepage and quick favorite modal window on home";						
					
						LogFactory.info("Favourite are not in user preferred lang on homepage ");
						
					
					}
				}
				else{

					flag = "Pass";
					result= "Favourites are not present on Favourites homepage Portlet";	
					
				}

				ReportFactory.reporterOutput(TCID, "Verify favorites for different languages ", expectedFavouriteLink,
						"The favorites should be displayed for the site irrespective of the languages selected on Favorites portlet on homepage ",
						result, flag);
		
					
			}catch(Exception e)
			{			
				String er = e.getMessage().substring(0, 160).toString().trim();

				ReportFactory.reporterOutput(TCID, "Verify favorites for different languages ", expectedFavouriteLink,
						"NA", er, flag);    	
			}
			
			//To click on dealerpath homepge
			homepagepath.click();
		}


// ****************************** E - N - D *****************************************************
				
		/*//Verify Favourites for different languages on department page.
		public static void verifyFavouritesDifferentLanguagesOnDepartment(String expectedFavouriteLink, String deptName) throws Throwable
		{
			String TCID= "WCM_10";		
			String flag = "Fail";
			String result= "Favourites are showing in user preferred languages on department ";
				
		try{
					//String expectedFavouriteLink = "GSA Emp OEM Link Test";
			// Click on Given Departmnet 
			GenericFactory.toClickonDeptOnFavourite(deptName);
			
			// To  Get Favourite Link Nmae		
			ArrayList<String> Favlinkname = ToGetFavouriteslink_Department();
			if(Favlinkname.contains(expectedFavouriteLink))
			{
				flag = "Pass";
				result= "Favourites are showing in user preferred languages on department ";
				
				LogFactory.info("Favourite is Present On  department ");
				
				ReportFactory.reporterOutput(TCID, "verify Favourites Different Languages On department ", "NA",
						"Favourites should show in user preferred languages on department", result, flag);			
				
			}
			else{
				//Fail
				flag = "Fail";
				result= "Favourites is not showing in user preferred languages on department ";
				
				LogFactory.info("Favourite are not in user preferred lang on department ");
				
				ReportFactory.reporterOutput(TCID, "verify Favourites Different Languages On department ", "NA",
						"Favourites should show in user preferred languages on department", result, flag);
			}
		}catch(Exception e)
			{			
				String er = e.getMessage().substring(0, 25).toString().trim();

				ReportFactory.reporterOutput(TCID, "verify Favourites Different Languages On department ", "NA",
						"NA", er, flag);    	
			}
			
			//To click on dealerpath homepge
			homepagepath.click();
		}
*/
// **************************************************** E - N - D **********************************************************************************************************
						
	
			 //Verify Favourites for non-preferred departments.

			// Verifying the Fav link on departmnet page with test data 
			public static void VerifyFavouriteForPreffredDept(String TCID, String deptName,String Favlink) throws Throwable
			{
				String flag = "Fail";
				String result= "Favourite is not showing for preffred departemnt ";
						
				try{
					
					// Click on Given Departmnet 
				if(GenericFactory.toClickonDeptOnFavourite(deptName))
					 {
					
					// To  Get Favourite Link Nmae		
					ArrayList<String> Favlinkname = ToGetFavouriteslink_Department();
					
					if(Favlinkname.contains(Favlink))
					{	
						flag = "Pass";
						result= "Favourite is showing for preffred departemnt ";
						LogFactory.info("Favourite is showing for preffred departemnt ");
						//pass
						ReportFactory.reporterOutput(TCID, "Verify Favourite For Preffred Department", Favlink,
								"Favourite should show only for preffred departemnt", result, flag);
						
					}
					else{
						flag = "Fail";
						result= "Favourite is not showing for preffred departemnt ";
						//Fail
						LogFactory.info("Favourite link is not  Present On  Departmnet ");
						ReportFactory.reporterOutput(TCID, "Verify Favourite For Preffred Department ", Favlink,
								"Favourite should show only for preffred departemnt", result, flag);
					}
					 }else
					 {
						    flag = "Fail";
							result= "Department is not Clickable";
							LogFactory.info(result);
							//pass
							ReportFactory.reporterOutput(TCID, "Verify Favourite For Preffred Department", Favlink,
									"Favourite should show only for preffred departemnt", result, flag); 
					 }
			}
				catch(Exception e)
				{			
						String er = e.getMessage().substring(0, 175).toString().trim();

						ReportFactory.reporterOutput(TCID, "Verify Favourite For Preffred Department ", "NA",
								"NA", er, flag);    	
					}
					
					//To click on dealerpath homepge
					homepagepath.click();					
	}

// **************************************************** E - N - D **********************************************************************************************************
				
				
				//  Verifying the Fav link on departmnet page with test data  for non Preffred dept
				public static void VerifyFavouriteForNonPreffredDept( String TCID, String deptName,String Favlink) throws Throwable
				{
					String flag = "Fail";
					String result= "Favourite is showing for non preffred departemnt ";
							
				try{					
					
					// Click on Given Departmnet 
					if(GenericFactory.toClickonDeptOnFavourite(deptName))
					{
					
					// To  Get Favourite Link Nmae
					
					ArrayList<String> Favlinkname = ToGetFavouriteslink_Department();
					if(!Favlinkname.contains(Favlink))
					{
						flag = "Pass";
						result= "Favourite is not showing for only non preffred departemnt ";
						LogFactory.info("Favourite is not showing for non preffred departemnt ");
						//pass
						ReportFactory.reporterOutput(TCID, "Verify favorites for non-preferred departments ", "deptName- "+deptName+" "+"Favlink- "+Favlink,
								"Favorites should not be displayed if the particular department is non-preferred department of any user ",
								result, flag);
						
						
					}
					else{
						//Fail
						flag = "Fail";
						result= "Favourite is  showing for non preffred departemnt ";
						//Fail
						LogFactory.info("Favourite is   Present On Departmnet ");
						ReportFactory.reporterOutput(TCID, "Verify favorites for non-preferred departments ", "deptName- "+deptName+" "+"Favlink- "+Favlink,
								"Favorites should not be displayed if the particular department is non-preferred department of any user",
								result, flag);					
					}
					}
					else
					{
						flag = "Fail";
						result= "Department is not Clickable ";
						//Fail
						LogFactory.info("Favourite is   Present On Departmnet ");
						ReportFactory.reporterOutput(TCID, "Verify favorites for non-preferred departments ", "deptName- "+deptName+" "+"Favlink- "+Favlink,
								"Favorites should not be displayed if the particular department is non-preferred department of any user",
								result, flag);		
						
					}
					
				}catch(Exception e)
				{			
					String er = e.getMessage().substring(0, 25).toString().trim();

					ReportFactory.reporterOutput(TCID, "Verify favorites for non-preferred departments ", "NA",
							"NA", er, flag);    	
				}
				
				//To click on dealerpath homepge
				homepagepath.click();
					
	}
				
				
// **************************************************** E - N - D **********************************************************************************************************
	/*			
	
		//Verify Favourite on changing the Product Type.			

		public static void VerifyFavouritesOnchangingProductTypeOnHomePage(String TCID, String ProductSegment, String FavouriteLink) throws Throwable
		{
			String flag = "Fail";
			String result= "Favorites is not filtered based on product types filters ";
			
		try{
				
			//To  unchcek product type  and select the given product segmnets
			ToSelectProductType(ProductSegment);
			
			ArrayList<String> LinknameList = ToGetFavouritesFolder_HomePage();
			if(LinknameList.contains(FavouriteLink))
			{
				LogFactory.info("Favourite link is displaying on Preffred product segmnets");
						
			}
			else
			{
				LogFactory.info("Favourite link is not displaying on Preffred product segmnets");						
			}
		
		ReportFactory.reporterOutput(TCID, "Verify favorite on changing the Product Type ", "NA",
					"Favorites should be filtered based on product types filters ",
					result, flag);
			
			
		}catch(Exception e)
		{			
			String er = e.getMessage().substring(0, 60).toString().trim();

			ReportFactory.reporterOutput(TCID, "Verify favorite on changing the Product Type ", "NA",
					"NA", er, flag);    	
		}
		
	}
		*/
		
	/*	public static void VerifyFavouritesOnchangingProductTypeOnHomePage(String userCountry,String wcmCountry,String UserProductName, String WcmProducts,String titleName,String strTCID) throws Throwable
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
                                      System.out.println("test");                           
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
*/

		
// **************************************************** E - N - D **********************************************************************************************************		
		
		//Verify copy Favourites for dealer user.
		
		public static void verifyCopyFavouriteForDealer(String TCID, String Dealeruserid) throws Throwable
		{
			String flag = "Fail";
			String result= "Favourite is not copied for Dealer ";
			String errorMessage ="";
					
		try{
				//To get all folder name and link names
			    GenericFactory.endImpersonateOrLogoutUser();
			    homepagepath.click();
			    
				ArrayList<String> copiedFolderName = ToGetFavouritesFolder_HomePage();
				ArrayList<String> copiedFavouriteLink = ToGetMarkedFavouriteslink_HomePage();
				
				
				//click on action btn
				favouiteActions.click();
			
				// Execute the Java Script for the element which we find out
				((JavascriptExecutor) BaseClass.wbDriver).executeScript("favActionsCopyFavourites();");
				
				//To send another employee id for copy
				
			//	BaseClass.wbDriver.findElement(By.xpath(".//*[@id='copy-from-user']")).sendKeys(k);
				Thread.sleep(1000);
				BaseClass.wbDriver.findElement(By.xpath(".//*[@id='copy-from-user']")).click();
				BaseClass.wbDriver.findElement(By.xpath(".//*[@id='copy-from-user']")).clear();
				BaseClass.wbDriver.findElement(By.xpath(".//*[@id='copy-from-user']")).sendKeys(Dealeruserid);
				
				//To clk on add btn to add RACF id
				BaseClass.wbDriver.findElement(By.xpath(".//*[@id='copy-Favourites-copy']")).click();
				Thread.sleep(5000);
				//To retrieve error msg

				List<WebElement> Errorlist = BaseClass.wbDriver.findElements(By.xpath(".//*[@id='copy-Favourites-modal']/div/div/div[2]/p"));
			    
				for(int i=0;i<Errorlist.size();i++){
					errorMessage=Errorlist.get(i).getText().toString().trim();
					if(errorMessage.length()>0)
					{
						break;
					}
				}
				
			     
				if(errorMessage.length()>0)
				{					
					// click on cancel btn
					Thread.sleep(5000);
					BaseClass.wbDriver.findElement(By.xpath(".//*[@id='copy-Favourites-modal']/div/div/div[3]/button[1]")).click();
					//report fail - invaild user
					  flag = "Fail";
					  result=errorMessage;
						LogFactory.info("Can not copy "+result);
					
				}
				else
				{
					//To impersonate
					homepagepath.click();
					GenericFactory.utilityMenuAdminClick();
					GenericFactory.impersonateUser(Dealeruserid);
					Thread.sleep(5000);
					//again to get all folder name and link names
					ArrayList<String> actualFolderName = ToGetFavouritesFolder_HomePage();
					ArrayList<String> actualFavouriteLink = ToGetMarkedFavouriteslink_HomePage();
					
					if((actualFolderName.contains(copiedFolderName)) &&(actualFavouriteLink.contains(copiedFavouriteLink)))
					{
						flag = "Pass";
						result="Copied Favourite folder to another employee successfully";
						LogFactory.info("Copied Favourite folder to another employee successfully ");
					}
					else {
						  flag = "Fail";
						  result="Does not able to copied Favourite folder to another employee";
						
						LogFactory.info("Does not able to copied Favourite folder to another employee ");
					}				
						
				}
				ReportFactory.reporterOutput(TCID, "Verify copy favorites for dealer user ", "NA",
						"Favorites from one dealer to another dealer should be copied only if those links are accessible (in terms of country tagging, product type tagging, additional RACF group tagging, department access) to the dealer who is copying them ",
						result, flag);	
		}catch(Exception e)
		{			
			String er = e.getMessage().substring(0, 60).toString().trim();

			ReportFactory.reporterOutput(TCID, "Verify copy favorites for dealer user ", "NA",
					"NA", er, flag);    	
		}
		if(!(errorMessage.length()>0))
		{
		GenericFactory.endImpersonateOrLogoutUser();
		}
		//To click on dealerpath homepge
		homepagepath.click();
		GenericFactory.utilityMenuAdminClick();
		GenericFactory.impersonateUser(BaseClass.strUserRACFID);
		
   }
		
		
		
// **************************************************** E - N - D **********************************************************************************************************
		
		
		//Verify copy Favourites for Employee/Authorized contingent user.

		public static void verifyCopyFavouriteForEmployee(String TCID, String EMPID) throws Throwable
		{

			//String TCID = "TC13_Favourite";
			String flag = "Fail";
			String result= "Favourite is not copied for employee ";
			String errorMessage ="";
					
		try{
				//To get all folder name and link names
			    GenericFactory.endImpersonateOrLogoutUser();
			    homepagepath.click();
			    
				ArrayList<String> copiedFolderName = ToGetFavouritesFolder_HomePage();
				ArrayList<String> copiedFavouriteLink = ToGetMarkedFavouriteslink_HomePage();
				
				
				//click on action btn
				favouiteActions.click();
				
			
				  
				// Execute the Java Script for the element which we find out
				((JavascriptExecutor) BaseClass.wbDriver).executeScript("favActionsCopyFavourites();");
				
				//To send another employee id for copy
				
			//	BaseClass.wbDriver.findElement(By.xpath(".//*[@id='copy-from-user']")).sendKeys(k);
				Thread.sleep(1000);
				BaseClass.wbDriver.findElement(By.xpath(".//*[@id='copy-from-user']")).click();
				BaseClass.wbDriver.findElement(By.xpath(".//*[@id='copy-from-user']")).clear();
				BaseClass.wbDriver.findElement(By.xpath(".//*[@id='copy-from-user']")).sendKeys(EMPID);
				
				//To clk on add btn to add RACF id
				BaseClass.wbDriver.findElement(By.xpath(".//*[@id='copy-Favourites-copy']")).click();
				Thread.sleep(5000);
				//To retrieve error msg

				List<WebElement> Errorlist = BaseClass.wbDriver.findElements(By.xpath(".//*[@id='copy-Favourites-modal']/div/div/div[2]/p"));
			    
				for(int i=0;i<Errorlist.size();i++){
					errorMessage=Errorlist.get(i).getText().toString().trim();
					if(errorMessage.length()>0)
					{
						break;
					}
				}
				
			     
				if(errorMessage.length()>0)
				{					
					// click on cancel btn
					Thread.sleep(5000);
					BaseClass.wbDriver.findElement(By.xpath(".//*[@id='copy-Favourites-modal']/div/div/div[3]/button[1]")).click();
					//report fail - invaild user
					  flag = "Fail";
					  result=errorMessage;
						LogFactory.info("Can not copy "+result);
					
				}
				else
				{
					//To impersonate
					homepagepath.click();
					GenericFactory.utilityMenuAdminClick();
					GenericFactory.impersonateUser(EMPID);
					Thread.sleep(5000);
					//again to get all folder name and link names
					ArrayList<String> actualFolderName = ToGetFavouritesFolder_HomePage();
					ArrayList<String> actualFavouriteLink = ToGetMarkedFavouriteslink_HomePage();
					
					if((actualFolderName.contains(copiedFolderName)) &&(actualFavouriteLink.contains(copiedFavouriteLink)))
					{
						flag = "Pass";
						result="Copied Favourite folder to another employee successfully";
						LogFactory.info("Copied Favourite folder to another employee successfully ");
					}
					else {
						  flag = "Fail";
						  result="Does not able to copied Favourite folder to another employee";
						
						LogFactory.info("Does not able to copied Favourite folder to another employee ");
					}				
						
				}
		
		ReportFactory.reporterOutput(TCID, "Verify copy favorites for Employee/Authorized contingent user ", "NA",
						"Favorites from one employee/authorized contingent user to another employee/authorized contingent user should be copied only if those links are accessible (in terms of additional RACF group tagging, Dealer Principal department access) to the employee/authorized contingent user who is copying them ",
						result, flag);	
		
		}catch(Exception e)
		{			
			String er = e.getMessage().substring(0, 60).toString().trim();

			ReportFactory.reporterOutput(TCID, "Verify copy favorites for Employee/Authorized contingent user ",
					"NA", "NA", er, flag);    	
		}
		if(!(errorMessage.length()>0))
		{
		GenericFactory.endImpersonateOrLogoutUser();
		}
		//To click on dealerpath homepge
		homepagepath.click();
		Thread.sleep(2000);
		GenericFactory.utilityMenuAdminClick();
		GenericFactory.impersonateUser(BaseClass.strUserRACFID);
				
	}	
		///****************************************************************************
		

		
// **************************************************** E - N - D **********************************************************************************************************
	
		
		
		
		
		
		public static void VerifyFavouritesOnchangingProductType(String departmentName, String LinkName, String WCMproductlist)
		
		{
			String Favlink="";
			String IconIsslected="";
			 ArrayList<String> QuciMarkedFavlink_Homepage= new ArrayList<String>();
			 ArrayList<String> MarkedFavlinkhomepage= new ArrayList<String>();
			 ArrayList<String> MarkedFavlinkonDepartmnet= new ArrayList<String>();
			 String ProducSegmnet="";
		 if( GenericFactory.toClickonDeptOnFavourite(departmentName))
		    {
			  List<WebElement> favouriteLinkList = favouiteLinkList;
			   List<WebElement> favouriteIconList = favouiteIcon;
			  ArrayList<String> Favlinklist = ToGetFavouriteslink_Department();
			    if(Favlinklist.size()>0)
			     {
			    	 for(int i=0;i<Favlinklist.size();i++)
					    {
			    		    Favlink=favouriteLinkList.get(i).getAttribute("textContent").trim();
			    		    
			    		    if(Favlink.equalsIgnoreCase(LinkName))
			    		    {
			    		        IconIsslected = favouriteIconList.get(i).getAttribute("class").toString().trim();
			    		    	 if(IconIsslected.contains("fav-star is-selected"))
			    			       {
			    			    	  LogFactory.info(LinkName +" is Already Marked so First Unmark the link and Create Marked fav");
			    			    	  favouriteIconList.get(i).click();
			    			       }
			    		    	   favouriteIconList.get(i).click();
			    		    	   favouriteSavebtn.click();
			    		    	   homepagepath.click();
			    		    	   // Link is marked as Fav
			    		    	   QuciMarkedFavlink_Homepage= ToGetMarkedFavouriteslinkQuickModal();
			    		    	   MarkedFavlinkhomepage= ToGetMarkedFavouriteslink_HomePage();
			    		    	   if(QuciMarkedFavlink_Homepage.contains(LinkName) &&  MarkedFavlinkhomepage.contains(LinkName))
			    		    	    {
			    		    		   LogFactory.info("Link is created as Fav and displaying on home page and present under Qucik model");
			    		    	    }
			    			         //
			    		    	   GenericFactory.checkAllProducts();
			    		    	  // String WCMproductlist = "jjd,jdjd";
			    		    	   if(WCMproductlist.contains(","))
			    		    	   {
			    		             String[] productarray = WCMproductlist.split(",");
			    		              for(int j=0;j<productarray.length;i++)
			    		               {
			    		            	ProducSegmnet=productarray[j];
			    		            	  GenericFactory.productUncheck(ProducSegmnet);
			    		              }
			    		    	     }
			    		    	   else{
			    		    		   ProducSegmnet=WCMproductlist;
			    		    		   GenericFactory.productUncheck(ProducSegmnet);
			    		    	   }
			    		   		   
			    		   		   MarkedFavlinkonDepartmnet = ToGetMarkedFavouriteslink_Department();
			    		   		   homepagepath.click();
			    		   		   QuciMarkedFavlink_Homepage= ToGetMarkedFavouriteslinkQuickModal();
			    		    	   MarkedFavlinkhomepage= ToGetMarkedFavouriteslink_HomePage();
			    		    	   //Check the link if Product is unchecked 
			    		    	   if(MarkedFavlinkonDepartmnet.contains(LinkName)&&QuciMarkedFavlink_Homepage.contains(LinkName)&&MarkedFavlinkhomepage.contains(LinkName))
			    		    	   {
			    		    		   //Fail
			    		    	   }
			    		    	   else
			    		    	   {
			    		    		   // Pass 
			    		    	   }
			    		    	   
			    		    }
					    }
			    	
			     }
			    else{
			    	// No Fav link present
			    }
			  
			 
		     }
			 else
			 {
				// Department is not clickable 
				 
			 }
			
		
					
				
		}
		
		
	// END
		
		
		
		
		public static void ToRemoveFavouriteLinkfromQuickmodalwindow(String TCID,String Markerfavlinktoremove) throws Throwable 
		{
			

			//ArrayList<String> QuickFolderNameList = new ArrayList<String>();
			String Markerfavlink="";
			String  flag = "Fail";
			String result= "Favourite is not removed and displaying on 'Favourite' Portlet on homepage,quick favourite modal window Home page and Department";
			ArrayList<String> Favlinklistquick = new ArrayList<String>();


			try {
				homepagepath.click();
				if(ToGetMarkedFavouriteslinkQuickModal().contains(Markerfavlinktoremove)&&Markerfavlinktoremove.length()>0)
				{	
					homepagepath.click();
					BaseClass.wbDriver.findElement(By.xpath(".//*[@id='js-favorites']")).click();
					Thread.sleep(2000);
					List<WebElement> QuickIconElmentList = BaseClass.wbDriver.findElements(By.xpath(".//*[@id='quick-favorites-target']/div/div/div/div/span"));
					List<WebElement> Quicklemetlist = BaseClass.wbDriver.findElements(By.xpath(".//*[@id='quick-favorites-target']/div/div/div/div"));
					Thread.sleep(2000);
					for(int i=0;i<Quicklemetlist.size();i++)
					{
						WebElement folderobj = Quicklemetlist.get(i);
						List<WebElement> linkobj = folderobj.findElements(By.tagName("a")) ;
						if((linkobj.size()>0))
						{

							Markerfavlink= folderobj.getText().toString().trim();
							if(Markerfavlink.equalsIgnoreCase(Markerfavlinktoremove))
							{
								Thread.sleep(2000);
								QuickIconElmentList.get(i).click();
								break;
							}
						} 

					}

					// 
					Thread.sleep(5000);
					ArrayList<String> Favlinklisthomepage = ToGetMarkedFavouriteslink_HomePage();
					Thread.sleep(2000);
					Favlinklistquick = ToGetMarkedFavouriteslinkQuickModal();

					if(Favlinklisthomepage.contains(Markerfavlinktoremove)&&Favlinklistquick.contains(Markerfavlinktoremove))
					{
						flag = "Fail";
						result= "Favourite is not removed and  displaying on 'Favourite' Portlet on homepage,quick favourite modal window Home page and Department";
						LogFactory.info("Favourite Link is not removed ");
						ReportFactory.reporterOutput(TCID, "verify Remove Favourite Link from QuickWindow Modal", Markerfavlinktoremove,
								"Favourite should get removed and should no more display on 'Favourite' Portlet on homepage,quick favourite modal window page", 
								result, flag); 

					}
					else
					{
						//
						flag = "Pass";
						result= "Favourite is removed and no more displaying on 'Favourite' Portlet on homepage,quick favourite modal window Home page and Department";
						LogFactory.info("Favourite Link is removed ");
						ReportFactory.reporterOutput(TCID, "verify Remove Favourite Link from QuickWindow Modal ", Markerfavlinktoremove,
								"Favourite should get removed and should no more display on 'Favourite' Portlet on homepage,quick favourite modal window page", 
								result, flag); 
						

					}

				}
				else
				{
					// Pass No fav link present on QuickmodalWindow
					flag = "Pass";
					result= Markerfavlinktoremove+" Link is not present  on QuickmodalWindow so can not remove";
					LogFactory.info(result);
					ReportFactory.reporterOutput(TCID, "verify Remove Favourite Link from QuickWindow Modal ", Markerfavlinktoremove,
							"Favourite should get removed and should no more display on 'Favourite' Portlet on homepage,quick favourite modal window page", 
							result, flag); 
				}
			} catch(Exception e)
			{			
				String er = e.getMessage().substring(0, 125).toString().trim();

				ReportFactory.reporterOutput(TCID, "verify Remove Favourite Link from QuickWindow Modal ", Markerfavlinktoremove,
						"Favourite should get removed and should no more display on 'Favourite' Portlet on homepage,quick favourite modal window page", 
						er, flag);   	
			}


	} 
			
		
		
		
		
		
		
		
		
		
		
	// ********************************************************************
		
		public static void ToRemoveFavLinkfromFavoritesPortlethomepage( String TCID,String Markerfavlinktoremove) throws Throwable
		{
			String Markerfavlink="";
			String  flag = "Fail";
			String result= "Favourite is not removed and displaying on 'Favourite' Portlet on homepage,quick favourite modal window Home page and Department";
			ArrayList<String> Favlinklistquick = new ArrayList<String>();
	
			
			try {
				homepagepath.click();
				if(ToGetMarkedFavouriteslink_HomePage().contains(Markerfavlinktoremove)&&Markerfavlinktoremove.length()>0)
				{
							
				List<WebElement> HomapageIconElmentList = BaseClass.wbDriver.findElements(By.xpath(".//*[@class='fav-link-body']/span"));
				List<WebElement> HomapageElemet = BaseClass.wbDriver.findElements(By.xpath(".//*[@class='fav-link-body']"));
				
				 for(int i=0;i<HomapageElemet.size();i++)
				 {
					 
					 WebElement folderobj = HomapageElemet.get(i);
					 List<WebElement> linkobj = folderobj.findElements(By.tagName("a")) ;
					 if((linkobj.size()>0))
					 {
						 
					     Markerfavlink= folderobj.getText().toString().trim();
					     if(Markerfavlink.equalsIgnoreCase(Markerfavlinktoremove))
					     {
					    	 HomapageIconElmentList.get(i).click();
					    	 break;
					     }
					 }  
				}
					 
					// 
					ArrayList<String> Favlinklisthomepage = ToGetMarkedFavouriteslink_HomePage();
					 Favlinklistquick = ToGetMarkedFavouriteslinkQuickModal();
					
					 if(Favlinklisthomepage.contains(Markerfavlinktoremove)&&Favlinklistquick.contains(Markerfavlinktoremove))
					   {
						    flag = "Fail";
							result= "Favourite is removed and no more displaying on 'Favourite' Portlet on homepage,quick favourite modal window Home page and Department";
							LogFactory.info("Favourite Link is not removed ");
							ReportFactory.reporterOutput(TCID, "verify Remove Favourite Link On FavoritesPortlethomepage ", Markerfavlinktoremove,
									"Favourite should get removed and should no more display on 'Favourite' Portlet on homepage,quick favourite modal window page", 
									result, flag);
				     	}
					   else
					    {
						    flag = "Pass";
							result= "Favourite is removed and no more displaying on 'Favourite' Portlet on homepage,quick favourite modal window Home page and Department";
							LogFactory.info("Favourite Link is removed ");
							ReportFactory.reporterOutput(TCID, "verify Remove Favourite Link On FavoritesPortlethomepage ", Markerfavlinktoremove,
									"Favourite should get removed and should no more display on 'Favourite' Portlet on homepage,quick favourite modal window page", 
									result, flag); 
	

					      }
				
				}
				else
				{
					flag = "Pass";
					result= Markerfavlinktoremove+" Link is not  present on FavoritesPortlethomepage so can not remove";
					LogFactory.info(result);
					ReportFactory.reporterOutput(TCID, "verify Remove Favourite Link on FavoritesPortlethomepage ", Markerfavlinktoremove,
							"Favourite should get removed and should no more display on 'Favourite' Portlet on homepage,quick favourite modal window page", 
							result, flag); 
				}
				
				
			} catch (Exception e) {
				String er = e.getMessage().substring(0, 125).toString().trim();

				ReportFactory.reporterOutput(TCID, "verify Remove Favourite Link from QuickWindow Modal ", Markerfavlinktoremove,
						"Favourite should get removed and should no more display on 'Favourite' Portlet on homepage,quick favourite modal window page", 
						er, flag);   
			}
			
		}
		
		
		
		//*********************************************************************
		
		

		
		/**
		 * @Functionality To verify the presence of Favorites header on homepage.
		 * @throws Throwable
		 */
		public static void verifyFavoritesHeaderPresent(String strTCID, String strExpectedHeaderMsg) throws Throwable {
			String strFlag = "Fail";
			try {
				String stractualheadermsg = wbelHeaderTitleFav.getText().toString();
				// System.out.println(stractualheadermsg);

				String strResult = "Favorites portlet with header text is not displayed.";

				if (ValidationFactory.isElementPresent(wbelHeaderTitleFav)) {

					if (stractualheadermsg.equalsIgnoreCase(strExpectedHeaderMsg)) {

						strFlag = "Pass";
						strResult = "Favorites portlet with header text is displayed.";
						LogFactory.info("Favorites portlet with header is displayed.");

					} else if (!stractualheadermsg.equalsIgnoreCase(strExpectedHeaderMsg)) {

						strFlag = "PASS";
						strResult = "Favourite Folders are created.";
						LogFactory.info("Verify Favourite Folders are created.");
						ReportFactory.reporterOutput(strTCID, "Verify Favourite Folders are created.",
								"Favourite Folders should be created.", "Favourite/Folders found.", strResult, strFlag);

					}

				}
				ReportFactory.reporterOutput(strTCID, "Verify Favorites portlet with header text is displayed on homepage.",
						strExpectedHeaderMsg, "Favorites portlet header should be " + strExpectedHeaderMsg, strResult,
						strFlag);

				if (strFlag.equalsIgnoreCase("FAIL")) {
					Assert.assertFalse(true);
				}
			}

			catch (Exception e) {

				ReportFactory.reporterOutput("Error_Favorites portlet", "Favorites portlet with header",
						"Favorites portlet with header", "NA", e.getMessage().toString(), strFlag);
			}
		}		

		
		
		/**
		 * @Functionality To verify No Favorites added yet message on homepage.
		 * @throws Throwable
		 *             exception.
		 */
		public static void verifyWhenNoFavouritePresent(String strTCID, String strExpectedNoFavMsg) throws Throwable {
			String strFlag = "Fail";
			try {
				String stractualnofavmsg = wbelNofavaddedFav.getText().toString();
				// System.out.println(stractualnofavmsg);

				String strResult = "Message 'No Favourites added yet' is not displayed.";
				if (ValidationFactory.isElementPresent(wbelNofavaddedFav)) {

					if (stractualnofavmsg.equalsIgnoreCase(strExpectedNoFavMsg)) {
						strFlag = "Pass";
						strResult = "Message 'No Favourites added yet' is displayed.";
						LogFactory.info("Verify 'No Favourites added yet.' message is displayed.");

					} else if (!stractualnofavmsg.equalsIgnoreCase(strExpectedNoFavMsg)) {

						strFlag = "PASS";
						strResult = "Favourite Folders are created.";
						LogFactory.info("Verify Favourite Folders are created.");
						ReportFactory.reporterOutput(strTCID, "Verify Favourite Folders are created.",
								"Favourite Folders should be created.", "Favourite/Folders found.", strResult, strFlag);

					}

					else {
						ReportFactory.reporterOutput(strTCID, "Verify No Favourites added yet message on homepage.",
								strExpectedNoFavMsg, "Message should be " + strExpectedNoFavMsg, strResult, strFlag);
					}

				}

			} catch (Exception e) {

				ReportFactory.reporterOutput(strTCID,"Verify 'No Favourites added yet' message on home page when No Favourites added.","NA", "NA", e.getMessage(), "Fail");
			}
		}	
			/**
			 * @Functionality : To verify favorite icon with message on homepage.
			 * @throws Throwable
			 */

			@SuppressWarnings("unused")
			public static void verifyFavouriteIconWhenNoFavouriteAdded(String TCID, String strExpectedIcon) throws Throwable {

				String flag = "Fail";
				try {
					String result = "Star icon with Message is not displayed.";

					if (lstActualFoldernames.size() > 0) {
						for (int j = 0; j < lstActualFoldernames.size(); j++) {
							String temp1 = lstActualFoldernames.get(j).getText().trim();
							System.out.println(temp1);
							flag = "PASS";
							result = "Currently Favourites are available so star icon with message is not showing.";
							break;
						}
						LogFactory.info("Verify Favourite Folders are created.");
					}

					else if (wbelMessageFav.isDisplayed())

					{
						String strActualIcon = wbelMessageFav.getText().toString();

						System.out.println(strActualIcon);

						if (strActualIcon.trim().equalsIgnoreCase(strExpectedIcon.trim())) {

							LogFactory.info("Verify Star icon with Message is displayed.");
							flag = "Pass";
							result = "Star icon with Message is displayed as expected..";
						}

					}

					ReportFactory.reporterOutput(TCID,
							"Verify Favorites Icon with message in Favourites portlet when no favourites", strExpectedIcon,
							"Star icon with Message should be  " + strExpectedIcon, result, flag);

				} catch (Exception e) {

					ReportFactory.reporterOutput(TCID, "Favorites Icon with Message.", "NA", "NA",e.getMessage().toString().substring(0, 125), flag);
				}

			}	

			
			public static void favouriteQuickLinkOnHomePage(String strTCID) throws Throwable {

				String strFlag = "FAIL";
				String strResult = "Favourite Quick link Icon is not displayed.";

				try {

					if (ValidationFactory.isElementPresent(wbelFavourSearch)) {
						LogFactory.info("Verify Favourite Quick link Icon is displayed.");
						strFlag = "PASS";
						strResult = "Favourite Quick link Icon is displayed.";

					}

					ReportFactory.reporterOutput(strTCID, "Verify Favourite Quick link Icon is displayed.", "Quick Link Icon",
							"Quick Link Icon should be displayed", strResult, strFlag);

				} catch (Exception e) {

					ReportFactory.reporterOutput("Error_Favorites Icon with Message", "Favorites quick link icon.",
							"Favorites quick link icon.", "NA", e.getMessage().toString(), strFlag);
				}

			}
		
			
			public static void homePageQuickLinkContent(String strTCID, String strExpectedHeaderName) throws Throwable {
				try {

					String strFlag = "FAIL";
					String strResult = "Favourite Quick link Icon with contents are not displayed.";

					if (ValidationFactory.isElementPresent(wbelFavourSearch)) {

						wbelFavourSearch.click();
						String strActualIconHeaderName = wbelQuickLinkHeader.getText().toString();
						System.out.println(strActualIconHeaderName);

						if (strActualIconHeaderName.equalsIgnoreCase(strExpectedHeaderName)) {

							LogFactory.info("Verify Favourite Quick link Icon with content is displayed.");
							strFlag = "PASS";
							strResult = "Favourite Quick link Icon with contents are displayed.";
						}

					}

					ReportFactory.reporterOutput(strTCID, "Verify Favourite Quick link Icon with content is displayed.",
							"Quick Link Icon with content", "Quick Link Icon with content is displayed", strResult, strFlag);

				} catch (Exception e) {

					ReportFactory.reporterOutput("Error_Favorites Icon with Message", "Favorites quick link icon.",
							"Favorites quick link icon.", "NA", e.getMessage().toString(), "Fail");
				}

			}
			
		
		
}
		 