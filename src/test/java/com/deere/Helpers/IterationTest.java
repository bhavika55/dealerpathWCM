package com.deere.Helpers;

import java.util.ArrayList;
import java.util.HashSet;
import java.util.LinkedHashSet;
import java.util.List;

import org.openqa.selenium.By;
import org.testng.TestNG;

import com.deere.PageFactory.Login_Page_POF;

/**
 * This method gets unique rackfID's from excel sheet and run the testng suite till the size of rackfID's
 * 
 * @author shrey.choudhary
 * @param strRACFID
 * @createdAt 05-06-2018
 * @throws IOException
 * 
 * @modifyBy 
 * @modifyAt 
 * 
 */
public class IterationTest extends BaseClass {
	
	public static void testNgSuite() throws Throwable {
	
		HashSet<String> racfIDs = ExcelFactory.uniqueRACFId();
		ArrayList<String> uniqueRackfId = new ArrayList<>(racfIDs);
		LogFactory.info(" Unique number of RACF ID's  >>>>>   " + uniqueRackfId.size() + " " + uniqueRackfId );
	
			try {
					for (int i = 0; i <uniqueRackfId.size(); i++) {
				
						if (ImpersonateUser.impersonateUserSuccess(uniqueRackfId.get(i))) {
						List<String> suites = new ArrayList<String>();
						strUserRACFID=uniqueRackfId.get(i);
						System.out.println("Reading excel data");
						ExcelFactory.getUserWCMContent();
						ExcelFactory.getDealersInfoFlag();
						loginPageFactory.verifyImpersonatedUser();

						if (flagAddtionalTestcases.equalsIgnoreCase("Yes")) {	
							LogFactory.info("Reading addtional testcases sheet for the dealer " + strUserRACFID) ;
							mapAddtionalTestcase= ExcelFactory.getUserAddtionalTestcases();
							suites.add(strWorkingDir + "\\TestNG_XML\\WCMTestingWithAddtionalTestcases.xml");
						}
						else {suites.add(strWorkingDir + "\\TestNG_XML\\WCMContentOnlyTesting.xml");}
						TestNG tng = new TestNG();
						tng.setTestSuites(suites);
						tng.run(); // run test suite				
						
						}
						else 
						{
							LogFactory.info("Entered Racf id is incorrect");
							//ReportFactory.tableEnd();
							GenericFactory.createHeaderSection("Error Impersoanate User " + BaseClass.strUserRACFID );
							ReportFactory.reporterOutput("Error Impersonate", "User unable to login",
									strUserRACFID, "User should redirect to Impersonate New User",
									"Reason could be : DealerPath Application is not available [AccountFlex is down] ", "Fail");
							 ValidationFactory.getElementIfPresent(By.xpath("//a[@id='endimpersonatelink']")).click();
							 ReportFactory.tableEnd();
						}
					
				}
				
				}catch (Exception e) {
					System.out.println(e.getMessage());
				
					LogFactory.info("Entered Racf id is incorrect");
			}
				
		
	}
}
