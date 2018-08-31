package com.deere.PageFactory;

import java.util.ArrayList;
import java.util.List;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.How;
import org.testng.Assert;
import org.testng.asserts.SoftAssert;

import com.deere.Helpers.GenericFactory;
import com.deere.Helpers.ReportFactory;
import com.deere.Helpers.ValidationFactory;

public class ProductSegment_POF {

	static WebDriver ProdDriver;
	static SoftAssert softAssert = new SoftAssert();

	public ProductSegment_POF(WebDriver driver) {
		this.ProdDriver = driver;

	}

	@FindBy(how = How.XPATH, using = "//div[@id='js-segments']")
	static WebElement wbelProductSegmentIcon;

	public static void checkOrderOfProductSegment(String strTestData, String strTCID) throws Throwable {

		String strFlag = "Fail";
		WebElement element = ValidationFactory.getElementIfPresent(By.xpath("//div[@id='js-segments']"));

		if (element != null) {
			{
				element.click();
			}
			List<String> listExpectedDataFromExcel = GenericFactory.splitString(strTestData, ",");
			List<String> listActualDataFromSite = GenericFactory.getCheckBoxValues();

			String strResult = "Product segments displayed are not in order:" + listActualDataFromSite.toString();

			try {
				if (listExpectedDataFromExcel.equals(listActualDataFromSite)) {
					strFlag = "Pass";
					strResult = "Product segments displayed are in order:" + listActualDataFromSite.toString();
					element.click();
				}
				ReportFactory.reporterOutput(strTCID,
						"Verify the product segments available in the product segments modal window ", strTestData,
						"Product segments should be in order " + listExpectedDataFromExcel.toString(), strResult,
						strFlag);

				/*
				 * if (strFlag.equalsIgnoreCase("FAIL")) { Assert.assertFalse(true); }
				 */

			}

			catch (Exception e) {

				ReportFactory.reporterOutput(strTCID,
						"Verify the product segments available in the product segments modal window ", "NA ",
						"Product segments should be in order " + listExpectedDataFromExcel.toString(),
						e.getMessage().toString(), strFlag);

			}

		}

	}
}
