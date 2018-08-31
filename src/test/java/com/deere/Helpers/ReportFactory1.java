package com.deere.Helpers;

import org.testng.Reporter;
import org.testng.annotations.Test;

public class ReportFactory1 
{
public static int ReprotDetails = 0;
	
	public static int stepnum=0;
	public static void reportDetailsHeader() 
	{
		Reporter.log("<table id='headerTable'; border=1px; style= width:100%;color=#808080>" + "<tr> <th width=20%>Browser Selected</th> <th width=20%> Environment Used</th></tr>");
		ReprotDetails = 1;
	}
	
	public static void addRoottable(){
		Reporter.log("<table id='headerTable'>");
	}
	
	public static void reportDetailsValues(String strPBrowserName, String strPEnvUsed)
	{
		String trstart = "<tr align=center>";
		String trend = "</tr>";
		String thstart = "<td width=15%>";
		String thend = "</td>";
		String tableend = "</table>";
		String BrowserName = strPBrowserName;
		String EnvUsed = strPEnvUsed;
		Reporter.log(trstart+thstart+ BrowserName+thend+thstart+EnvUsed + thend+trend+tableend);
	}
	public static void reportSectionName(String strPSectionName)
	{
		String TableStart = "<table border=1px; style= width:100%;color=#808080;body bgcolor=#F5DEB3>";
		
		String trstart = "<tr align=center>";
		String trend = "</tr>";
		String thstart = "<th width=15% >";
		String thend = "</th>";
		String SectionName = strPSectionName;
		String TableEnd = "</table>";
		
		Reporter.log(TableStart+trstart+thstart+SectionName+thend+trend+TableEnd);
	}
	
	@Test
	public static void reporterOutputHeader() throws InterruptedException 
	{
		Thread.sleep(1000);
	//	stepnum=0;
		Reporter.log("<table border=1px; style= width:14%;color=#808080> "
		+ " <tr font-family:Georgia, Garamond, Serif><th width=14%;body bgcolor=#BDBDBD>Testcase ID</th> <th width=14%;body bgcolor=#BDBDBD>Testcase Description</th> <th width=14%;body bgcolor=#BDBDBD>Input Value</th> <th width=14%;body bgcolor=#BDBDBD>Expected Value</th> <th width=14% ;body bgcolor=#BDBDBD>Actual Value</th> "
				+ "<th width=14% ;body bgcolor=#BDBDBD>Result</th></tr>"+"<th width=14% ;body bgcolor=#BDBDBD>Snapshot</th></tr>");

	}

	public static void reporterOutput(String strPTestCaseID,String strPStepDescription,
			String strPInputValue, String strPExpectedValue, String strPActualValue,
			String strPResult, String strFilePath) throws Throwable {
		// String TableHeader = "<table border=1 style= width:100%> " +
		// "<tr> <th>Step Description</th> <th>Input Value</th> <th>Expected Value</th> <th>Actual Value</th> <th>Result</th></tr>";
		Thread.sleep(1000);
//		stepnum++;
	//	String Stepnumber="STEP "+stepnum+"";
		//String TestcaseID = Stepnumber;
		String TestcaseID = strPTestCaseID;
		String StepDescription = strPStepDescription;
		String InputValue = strPInputValue;
		String ExpectedValue = strPExpectedValue;
		String ActualValue = strPActualValue;
		String Result = strPResult;
		String screenshotPath="<a href="+strFilePath+" target=\"_blank\">Link</a>";
		

		String trstart = "<tr id=" + strPTestCaseID + ">";
		String trend = "</tr>";
		// String thstart = "<td width=20% id="+Id+">";
		String thstart = "<td width=20%>";
		// String thstart = "<td width=20% id="+ +">";
		String thend = "</td>";

		if (Result.equalsIgnoreCase("Pass")) {
			Result = "<TD align=center BGCOLOR= #00ff00 > Pass";
		}
		if (Result.equalsIgnoreCase("Fail")) {
			Result = "<TD align=center BGCOLOR= #A71313 > Fail";
		}

		Reporter.log(trstart +  TestcaseID + thend + thstart + StepDescription + thend + thstart
				+ InputValue + thend + thstart + ExpectedValue + thend	+ thstart + ActualValue + thend + Result + thend+thstart+ screenshotPath+thend);
	}

	public static void tableEnd() 
	{
		Reporter.log("</table>");
	}
	
	public static void exportToExcel() 
	{
		//Reporter.log("<button type=\"button\" onClick=\"fnExcelReport();\">ExportToExcel</button>");
		Reporter.log("<button type=\"button\" onClick=\"fnExcelReport();\">ExportToExcel</button>"
				+ "<script type=\"text/javascript\">function fnExcelReport(){ console.log(\"test\"); "
				+" var tab_text='<table border=\"2px\"><tr bgcolor=\"#87AFC6\">'; "
				+" var textRange; var j=0; "
				+" var tab = document.getElementById('headerTable'); "
				+" for(j = 0 ; j < tab.rows.length ; j++)  { tab_text=tab_text+tab.rows[j].innerHTML+\"</tr>\";}"
				+" tab_text=tab_text+'</table>'; "
				+" tab_text= tab_text.replace(/<A[^>]*>|<\\/A>/g, \"\");"
				+" tab_text= tab_text.replace(/<img[^>]*>/gi,\"\");"
				+" tab_text= tab_text.replace(/<input[^>]*>|<\\/input>/gi, \"\"); "
				+" var ua = window.navigator.userAgent;"
				+" var msie = ua.indexOf(\"MSIE\"); "
				+" if (msie > 0 || !!navigator.userAgent.match(/Trident.*rv\\:11\\./)){"
				+" txtArea1.document.open(\"txt/html\",\"replace\");"
				+" txtArea1.document.write(tab_text);"
				+" txtArea1.document.close();"
				+" txtArea1.focus();"
				+" sa=txtArea1.document.execCommand(\"SaveAs\",true,\"Thanks.xls\");"
				+" } else"
				+" sa = window.open('data:application/vnd.ms-excel,' + encodeURIComponent(tab_text));"
				+" return (sa); } "				
				+" </script> ");
		
		Reporter.log(" <iframe id=\"`\" style=\"display:none\"></iframe>");
		
		//Reporter.log("<button type=\"button\">ExportToExcel</button><script type='text/javascript'>function ExportToExcel(mytblId){var htmltable= document.getElementById('my-table-id');var html = htmltable.outerHTML;window.open('data:application/vnd.ms-excel,' + encodeURIComponent(html));}</script>");
	}
	
	
	
	
	/*function fnExcelReport()
{
    var tab_text="<table border='2px'><tr bgcolor='#87AFC6'>";
    var textRange; var j=0;
    tab = document.getElementById('headerTable'); // id of table

    for(j = 0 ; j < tab.rows.length ; j++)  { tab_text=tab_text+tab.rows[j].innerHTML+"</tr>";
        //tab_text=tab_text+"</tr>";
    }

    tab_text=tab_text+"</table>";
    tab_text= tab_text.replace(/<A[^>]*>|<\/A>/g, "");//remove if u want links in your table
    tab_text= tab_text.replace(/<img[^>]*>/gi,""); // remove if u want images in your table
    tab_text= tab_text.replace(/<input[^>]*>|<\/input>/gi, ""); // reomves input params

    var ua = window.navigator.userAgent;
    var msie = ua.indexOf("MSIE "); 

    if (msie > 0 || !!navigator.userAgent.match(/Trident.*rv\:11\./))      // If Internet Explorer
    {
        txtArea1.document.open("txt/html","replace");
        txtArea1.document.write(tab_text);
        txtArea1.document.close();
        txtArea1.focus(); 
        sa=txtArea1.document.execCommand("SaveAs",true,"Say Thanks to Sumit.xls");
    }  
    else                 //other browser not tested on IE 11
        sa = window.open('data:application/vnd.ms-excel,' + encodeURIComponent(tab_text));  

    return (sa);
}*/
	
	

}