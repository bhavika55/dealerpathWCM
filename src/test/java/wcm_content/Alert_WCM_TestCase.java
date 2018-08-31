package wcm_content;


import org.testng.annotations.Test;

import com.deere.Helpers.BaseClass;



public class Alert_WCM_TestCase extends BaseClass{
	
	
		
	@Test
	public static void AlertWCM() throws Throwable {
		try
		{
			loginPageFactory.navigateToWCM();
		}
		
		catch(Exception e)
		{
			
			System.out.println("WCM not navigated properly"+e.getMessage().toString());
		}
	}

}
