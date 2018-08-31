
package com.deere.Helpers;

import java.util.ArrayList;
import java.util.List;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.How;
import org.testng.annotations.Test;

import com.deere.Helpers.GenericFactory;

class MyPrefDept_POF {

	static WebDriver locDriver;

	public MyPrefDept_POF(WebDriver driver) {

		this.locDriver = driver;
	}

	@FindBy(how = How.XPATH, using = "//input[@type='checkbox' and @checked='true']")
	static List<WebElement> checkedValues;

	@FindBy(how = How.XPATH, using = "//*[@class='group-value checkbox-value']//label[@class='click-target-only']//input[@type='checkbox']")
	static List<WebElement> allCheck;

	public static void openMyPrefModal() throws InterruptedException {
		GenericFactory.utilityMenuMyPreferenceClick();
	}

	public static void toggleCheck(List<WebElement> allCheckboxes, List<String> checklist) {
		for (int i = 0; i < allCheckboxes.size(); i++) {
			if (checklist.contains(allCheckboxes.get(i).getAttribute("value"))) {
				allCheckboxes.get(i).click();
			}
		}
	}

	public static void saveMyPref() {
		locDriver.findElement(By.xpath("//*[@id='preference-save']")).click();
	}

	public static void myPrefRun() throws InterruptedException {
		openMyPrefModal();
		Thread.sleep(3000);

		// WebElement
		// Wbel=locDriver.findElement(By.xpath("//*[@id='myModalLabel']"));
		// WaitFactory.waitForElement(Wbel,locDriver);
		List<String> checklist = new ArrayList<>();
		List<WebElement> allCheckboxes = allCheck;

		for (int i = 0; i < allCheckboxes.size(); i++) {
			if (allCheckboxes.get(i).getAttribute("checked") == null) {
				checklist.add(allCheckboxes.get(i).getAttribute("value"));
			}
		}

		toggleCheck(allCheckboxes, checklist);
		saveMyPref();
		Thread.sleep(10000);
		// WebElement
		// Wbel1=locDriver.findElement(By.xpath("//*[@id='js-user-info']/span[2]"));
		// WaitFactory.waitForElement(Wbel1,locDriver);
		openMyPrefModal();
		Thread.sleep(3000);
		// WaitFactory.waitForElement(Wbel,locDriver);
		toggleCheck(allCheckboxes, checklist);
		saveMyPref();

	}
}

//************************************************************************************************************************
@Test
public class MyPrefRun {


	public static void MyPrefCheckboxRun() throws InterruptedException {
		MyPrefDept_POF.myPrefRun();
	}

}
