package com.deere.TestCasesFactory;

import org.openqa.selenium.WebDriver;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import com.deere.Helpers.EndImpersonateUser;
import com.deere.Helpers.GenericFactory;
import com.deere.Helpers.ReportFactory;

public class EndImpersonate_TestCase {
	
	WebDriver driver;

	@BeforeClass
	public void getReportHeader() throws InterruptedException {
		ReportFactory.tableEnd();
	//	GenericFactory.createHeaderSection("Home_Page");
		
	}
	
	@Test
	public void endImpersonateUser() {
		System.out.println("@@@@@@@@@@@@@@@@@@@@@@@@");
		new EndImpersonateUser();
		EndImpersonateUser.endImpersonateOrLogoutUser();
	}

}
