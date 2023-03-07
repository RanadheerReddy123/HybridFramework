package driverFactory;

import org.testng.Reporter;
import org.testng.annotations.Test;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

import commonFunctions.FunctionLibrary;
import config.AppUtil;
import utilities.ExcelFileUtil;

public class AppTest extends AppUtil{
	String inputpath ="C:/RR/eclipse-workspace/Hybrid_FrameWork/DataTables/DataEngine.xlsx";
	String outputpath ="C:/RR/eclipse-workspace/Hybrid_FrameWork/TestResults/HybridResults.xlsx";
	String TCSheet ="MasterTestCases";
	String TSSheet ="TestSteps";
	ExtentReports report;
	ExtentTest test;
	@Test
	public void startTest()throws Throwable
	{
		//define path for HTML
		report= new ExtentReports("./Reports/Hybrid.HTML");
		boolean res=false;
		String tcres="";
		//access excel methods
		ExcelFileUtil xl = new ExcelFileUtil(inputpath);
		//count no of rows in master test cases sheet
		int TCCount =xl.rowCount(TCSheet);
		//count no of rows in Test step sheet
		int TSCount =xl.rowCount(TSSheet);
		Reporter.log("No of rows in Master TCS sheet::"+TCCount+"     "+"No of rows in Teststep sheet::"+TSCount,true);
		//iterate all rows in master test case sheet
		for(int i=1;i<=TCCount;i++)
		{
			//read module name cell data
			String ModuleName=xl.getCellData(TCSheet, i, 1);
			//start test case
			test=report.startTest(ModuleName);
			//read execution mode cell
			String executionMode =xl.getCellData(TCSheet, i, 2);
			if(executionMode.equalsIgnoreCase("Y"))
			{
				//read tcid cell from TCSheet
				String tcid =xl.getCellData(TCSheet, i, 0);
				//iterate all rows in TSSheet
				for(int j=1;j<=TSCount;j++)
				{
					//read description cell
					String Description =xl.getCellData(TSSheet, j, 2);
					//read tsid cell from Tssheet
					String tsid =xl.getCellData(TSSheet, j, 0);
					if(tcid.equalsIgnoreCase(tsid))
					{
						//read keyword cell tssheet
						String keyword =xl.getCellData(TSSheet, j, 3);
						if(keyword.equalsIgnoreCase("AdminLogin"))
						{
							String para1 =xl.getCellData(TSSheet, j, 5);
							String para2 =xl.getCellData(TSSheet, j, 6);
							//call login method
							res=FunctionLibrary.verifyLogin(para1, para2);
							test.log(LogStatus.INFO, Description);

						}
						else if(keyword.equalsIgnoreCase("NewBranch"))
						{
							String para1 =xl.getCellData(TSSheet, j, 5);
							String para2 =xl.getCellData(TSSheet, j, 6);
							String para3 =xl.getCellData(TSSheet, j, 7);
							String para4 =xl.getCellData(TSSheet, j, 8);
							String para5 =xl.getCellData(TSSheet, j, 9);
							String para6 =xl.getCellData(TSSheet, j, 10);
							String para7 =xl.getCellData(TSSheet, j, 11);
							String para8 =xl.getCellData(TSSheet, j, 12);
							String para9 =xl.getCellData(TSSheet, j, 13);
							FunctionLibrary.clickBranches();
							res =FunctionLibrary.verifyBranchCreation(para1, para2, para3, para4, para5, para6, para7, para8, para9);
							test.log(LogStatus.INFO, Description);

						}
						else if(keyword.equalsIgnoreCase("BranchUpdate"))
						{
							String para1 =xl.getCellData(TSSheet, j, 5);
							String para2 =xl.getCellData(TSSheet, j, 6);
							String para5 =xl.getCellData(TSSheet, j, 9);
							String para6 =xl.getCellData(TSSheet, j, 10);
							FunctionLibrary.clickBranches();
							res =FunctionLibrary.verifyBranchUpdation(para1, para2, para5, para6);
							test.log(LogStatus.INFO, Description);
						}
						else if(keyword.equalsIgnoreCase("AdminLogout"))
						{
							res =FunctionLibrary.verifyLogout();
							test.log(LogStatus.INFO, Description);
						}

						String tsres ="";

						if(res)
						{
							//if res is true write pass into status cell
							tsres="Pass";
							xl.setCellData(TSSheet, j, 4, tsres, outputpath);
							test.log(LogStatus.PASS, Description);
						}
						else
						{
							//if res is false  write Fail into status cell
							tsres="Fail";
							xl.setCellData(TSSheet, j, 4, tsres, outputpath);
							test.log(LogStatus.FAIL, Description);
						}
						tcres=tsres;

					}
					report.endTest(test);
					report.flush();
				}
				//write tcres into TCSheet
				xl.setCellData(TCSheet, i, 3, tcres, outputpath);
			}
			else
			{
				//write as blocked which tc is flagged to N in TCSheet
				xl.setCellData(TCSheet, i, 3, "Blocked", outputpath);
			}
		}
	}
}













