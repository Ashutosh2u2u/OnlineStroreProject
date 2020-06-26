package testCases;

import org.apache.log4j.xml.DOMConfigurator;
import org.openqa.selenium.WebDriver;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Test;

import pageObjects.BaseClass;
import utility.Constant;
import utility.ExcelUtils;
import utility.Log;
import utility.Utils;
import appModules.CheckOut_Action;
import appModules.Confirmation_Action;
import appModules.PaymentDetails_Action;
import appModules.ProductSelect_Action;
import appModules.SignIn_Action;
import appModules.Verification_Action;

public class Framework_003{
	public WebDriver driver;
	private String sTestCaseName;
	private int iTestCaseRow;

	@BeforeMethod
	public void beforeMethod() throws Exception {

		DOMConfigurator.configure("log4j.xml");		
		sTestCaseName = this.toString();		
		sTestCaseName = Utils.getTestCaseName(this.toString());		
		Log.startTestCase(sTestCaseName);
		System.out.println(sTestCaseName);
		ExcelUtils.setExcelFile(Constant.Path_TestData + Constant.File_TestData,"Sheet1");
		iTestCaseRow = ExcelUtils.getRowContains(sTestCaseName,Constant.Col_TestCaseName);		
		driver = Utils.OpenBrowser(iTestCaseRow);
		new BaseClass(driver);  
	}

	// This is the starting of the Main Test Case
	@Test
	public void main() throws Exception {

		try{
			SignIn_Action.Execute(iTestCaseRow);

			///ProductSelect_Action.productType(iTestCaseRow);
			///ProductSelect_Action.productNumber(iTestCaseRow);
			///CheckOut_Action.Execute();
			//PaymentDetails_Action.execute(iTestCaseRow);
			//Confirmation_Action.Execute();
			//Verification_Action.Execute();

			System.out.println(BaseClass.bResult);
			if(BaseClass.bResult==true)
			{
				ExcelUtils.setCellData("Pass", iTestCaseRow, Constant.Col_Result);
			}
			else
			{
				throw new Exception("Test Case Failed because of Verification");
			}
		}
		catch (Exception e)
		{
			ExcelUtils.setCellData("Fail", iTestCaseRow, Constant.Col_Result);
			Utils.takeScreenshot(driver, sTestCaseName);
			Log.error(e.getMessage());
			throw (e);
		}
	}

	@AfterMethod
	public void afterMethod() {
		Log.endTestCase(sTestCaseName);
		driver.close();
	}
}
