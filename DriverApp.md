# Selenium-Automation
#DriverApp.java

package lib;

import org.testng.annotations.Test;
import org.testng.annotations.Test;
import org.testng.annotations.Test;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.IOException;
import java.lang.reflect.Method;
import java.net.MalformedURLException;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.List;
import java.util.Properties;


//import org.sikuli.basics.Settings;
//import org.sikuli.script.Screen;
//import org.sikuli.basics.Settings;
/*import org.sikuli.script.Screen;*/
import org.testng.annotations.AfterSuite;
import org.testng.annotations.BeforeSuite;
import org.testng.annotations.Test;
import org.apache.pdfbox.pdfparser.PDFParser;
import org.openqa.selenium.WebDriver;

import com.thoughtworks.selenium.DefaultSelenium;

import lib.TestReports;
import lib.SeleniumLib;
import lib.TestUtil;
import lib.Xlfile_Reader;

public class DriverApp {
	public static Properties GLOBAL;
	public static Properties CONFIG;
	public static Properties Objects;
	public static Properties CommonData;
	public static Xlfile_Reader controller;
	public static Xlfile_Reader testData=null;
	public static Xlfile_Reader testStep;
	public static String keyword;
	public static WebDriver driver;
	public static DefaultSelenium selenium=null;
	public static String object;
	public static String currentFTID = null;
	public static String currentTCID = null;
	public static String currentFTName;
	public static String currentTCName;
	public static String currentTSID;
	public static String stepDescription;
	public static String proceedOnFail;
	public static String tsEnvironment;
	public static String data_column_name;
	public static String tsData;
	public static List<String> listTsData;
	public static Calendar cal = new GregorianCalendar();
	public static  int month = cal.get(Calendar.MONTH);
	public static int year = cal.get(Calendar.YEAR);
	public static  int sec =cal.get(Calendar.SECOND);
	public static  int min =cal.get(Calendar.MINUTE);
	public static  int date = cal.get(Calendar.DATE);
	public static  int day =cal.get(Calendar.HOUR_OF_DAY);
	public static String strDate;
	public static String tsResult;
	public static String tcResult;
	public static String mailscreenshotpath;
	public static String curTestSuite;
	public static String suiteStartTime;
	public static String suiteEndTime;
	public static String featureStartTime;
	public static String suitePath;
	public static String reportPath;
	public static String reportCurrentPath;
	public static String featurelistFilePath;
	public static int FIDCount;
	public static int FCount=0;
	public static String currentLogName;
	public static String PASS;
	public static String FAIL;
	public static String separator;
	public static String StartDate;
	public static String StartDate1;
	public static  PDFParser parser=null;
	
	public static boolean isSignedIn = false;
	public static String startTime=null;
	public static FileInputStream fs;
	public static final String PROCESS_INTERNET_EXPLORER = "iexplore.exe";
	public static String getCurrentTimeStamp()
    { 
          SimpleDateFormat CurrentDate = new SimpleDateFormat("MM/dd/yyyy"+" "+"hh:mm:ss"+" "+"a");
          Date now = new Date(); 
         String CDate = CurrentDate.format(now); 
          
          return CDate; 
    }
	public static String getCurrentDate()
    { 
          SimpleDateFormat CurrentDate = new SimpleDateFormat("MM/dd/yyyy");
          Date now = new Date(); 
          String CDate = CurrentDate.format(now); 
          
          return CDate; 
    }


   	@BeforeSuite
	public void startTesting() throws Exception
	{
        try 
        {
        	 try {
					Runtime.getRuntime().exec("taskkill /F /IM IEDriverServer.exe");
					 System.out.println("taskkill");
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
			   
        strDate=getCurrentTimeStamp();
		System.out.println("date time stamp :"+strDate);
		
        //Loading Config File
		CONFIG = new Properties();
		FileInputStream fs = new FileInputStream(System.getProperty("user.dir")+"\\src\\test\\java\\config\\config.properties");
		CONFIG.load(fs);
		
		// LOAD OR properties File
		Objects = new Properties();
		fs = new FileInputStream(System.getProperty("user.dir")+"\\src\\test\\java\\config\\object.properties");
		Objects.load(fs);

		//Loading Global File
		GLOBAL = new Properties();
		fs = new FileInputStream(System.getProperty("user.dir")+"\\src\\test\\java\\config\\global.properties");
		GLOBAL.load(fs);
		
		// Initialize data files
		controller= new Xlfile_Reader(System.getProperty("user.dir")+ GLOBAL.getProperty("coreFile"));
		testData  =  new Xlfile_Reader(System.getProperty("user.dir")+ GLOBAL.getProperty("testdataFile"));

		// Create a run-time folder for execution report
		reportPath = System.getProperty("user.dir") + GLOBAL.getProperty("reportPath");
		StartDate=strDate.replaceAll("[:\\/ ]", "-");
		//StartDate1=StartDate.replaceAll("\\.", "-");
		
		
		
		TestUtil.createDirectory(reportPath, StartDate);
		reportCurrentPath = reportPath + StartDate + "\\";
		System.out.println(reportCurrentPath);
		// Initial global variable for test result
		PASS = GLOBAL.getProperty("resultPass");
		FAIL = GLOBAL.getProperty("resultFail");
		
		// Initial common global variables
		separator = CONFIG.getProperty("data_list_separator");
	}
	catch (Exception e) 
	{
		System.err.println("Failed with reason: " + e.getMessage());
		e.printStackTrace();
	}
	}
	
	@Test

	public void testApp(){//throws MalformedURLException {
		
		
		curTestSuite = CONFIG.getProperty("currentTestSuite");
		suiteStartTime = TestUtil.now("dd.MMMMM.yyyy hh.mm.ss aaa");
		featurelistFilePath = reportCurrentPath + "featurelist"+StartDate+GLOBAL.getProperty("reportFileExt");
		TestReports.startFeature(featurelistFilePath);
		FIDCount=controller.getRowCount(curTestSuite);
		// Loop FT in Suite1
		for(int ftid = 2; ftid <= FIDCount; ftid++)
		{
			// Check the FT Runmode
			if(controller.getCellData(curTestSuite, "Runmode", ftid).equalsIgnoreCase("Y")) 
			{
				featureStartTime = TestUtil.now("dd.MMMMM.yyyy hh.mm.ss aaa");
				FCount++;
				currentFTID = controller.getCellData(curTestSuite, "FTID", ftid);
				currentFTName = controller.getCellData(curTestSuite, "FT_Name", ftid);
				System.out.println("Executing Feature ID 12345= " + currentFTID);
			
				// Initialize environment for the feature prior to execution
				DriverLib.featureEnter();
				
				String testcaselistFilePath = reportCurrentPath + currentFTID + "_feature_index_"+StartDate+GLOBAL.getProperty("reportFileExt");
				TestReports.startTestCase(testcaselistFilePath, currentFTID);
				
				int countPassPerFeature = 0;
				int countFailPerFeature = 0;
				
				// Loop TC in currentFTID
				for (int tcid = 2; tcid <= controller.getRowCount(currentFTID); tcid++)
				{
					// Check the TC Runmode
					if(controller.getCellData(currentFTID, "Runmode", tcid).equalsIgnoreCase("Y")) 
					{
						currentTCID = controller.getCellData(currentFTID, "TCID", tcid);
						currentTCName = controller.getCellData(currentFTID, "TC_Name", tcid);
						System.out.println("Executing Testcase ID = " + currentTCID);
						
						proceedOnFail=controller.getCellData(currentFTID, "ProceedOnFail", tcid);
						String ftPath = GLOBAL.getProperty("testcaseDir") + currentFTID + "\\";
						testStep = new Xlfile_Reader(System.getProperty("user.dir") + ftPath + currentTCID + ".xlsx");
						System.out.println(ftPath);
						startTime=TestUtil.now("dd.MMMMM.yyyy hh.mm.ss aaa");
				
						// Create log file for each test case
						currentLogName = currentFTID + "_" + currentTCID + StartDate + GLOBAL.getProperty("logFileExt");
						TestReports.createExecutionLog(reportCurrentPath, currentLogName);
						
						// Initialize environment for the test case prior to execution
						DriverLib.testcaseEnter();

						for(int tsid = 2; tsid <= testStep.getRowCount(currentTCID); tsid++) 
						{
							String strTSID = testStep.getCellData(currentTCID, "TSID", tsid);
							if(strTSID != "") 
							{
								System.out.println("Executing Test Step = " + testStep.getCellData(currentTCID, "Description", tsid));
								try
								{
									// Invoke a relevant method(keyword) to execute
									keyword=testStep.getCellData(currentTCID, "Keyword", tsid);
									object=testStep.getCellData(currentTCID, "Object", tsid);
									currentTSID=testStep.getCellData(currentTCID, "TSID", tsid);
									stepDescription=testStep.getCellData(currentTCID, "Description", tsid);
									System.out.println(stepDescription);
									tsData=testStep.getCellData(currentTCID, "Data", tsid);
									/*if(tsData.contains(",")) {
										listTsData = Arrays.asList(tsData.split("\\s*,\\s*"));
									}*/
									
									Method method= KeywordsApp.class.getMethod(keyword);
									tsResult = (String)method.invoke(method);
									
								}
								catch(Throwable t)
								{
								/*	System.out.println("Error came from " 
											+ "keyword= " + keyword
											+ " , object= " + object 
											+ " , testData= " + testData.getCellData(currentFTID, tsData, currentTCID));
									tsResult = FAIL;*/
									driver.quit();
									
								}
								System.out.println("Executing keyword= " + keyword + " , tsResult= " + tsResult);
	
								String screenshotFile = "";						
								if(driver != null){
									String screenshotDir = reportCurrentPath;
									if(tsResult.equalsIgnoreCase(FAIL) ||  keyword.equalsIgnoreCase("assertIntegerValue") 
											|| keyword.equalsIgnoreCase("calculate_TP_premium")|| keyword.equalsIgnoreCase("calculatepremium")
											|| "getPolicyNo".equalsIgnoreCase(keyword)) {
										screenshotFile = SeleniumLib.captureWebScreenShot(driver, screenshotDir, currentFTID, currentTCID, strTSID);
									}
									//screenshotFile = SeleniumLib.captureWebScreenShot(driver, screenshotDir, currentFTID, currentTCID, strTSID);
								}
								TestReports.addKeyword(stepDescription, keyword, tsResult, screenshotFile);
							}
							if(tsResult.equalsIgnoreCase(FAIL))
							{
							driver.quit();	
							break;
							}
							
						} //end for: loop in CurrentTCID
						tcResult = tsResult;
						System.out.println("Reporting test case= " + tcResult);
						
						// Cleanup environment after TC execution
						DriverLib.testcaseExit();
						
						// Count the total of Pass, Fail and Error for each feature
						if(tcResult.equalsIgnoreCase(PASS)){
							countPassPerFeature++;
						}else if(tcResult.equalsIgnoreCase(FAIL)) {
							countFailPerFeature++;
							//driver.close();
							//Runtime.getRuntime().exec("cmd /c Taskkill /IM IEDriverServer.exe /F");
						}
							
						// Report TS result as Pass or Fail or Error
						TestReports.addTestCase(currentTCID, startTime, TestUtil.now("dd.MMMMM.yyyy hh.mm.ss aaa"),	tcResult);
						
						// Check TC ProceedOnFail here to decide continue or stop
						if(tcResult.equalsIgnoreCase(FAIL) && proceedOnFail.equalsIgnoreCase("N"))
						{
							// Break the FOR loop around of current feature, and run the next feature
							System.out.println("Stop execution for this FT because the TC is FAIL & proceedOnFail=N");
							break; 
						}
					} // end if: TC Runmode
				} // end for: Loop TC in currentFTID
				
				// Cleanup environment for feature after execution
				DriverLib.featureExit();
				
				TestReports.endFeature(testcaselistFilePath, countPassPerFeature, countFailPerFeature);
			} // end: Check the FT Runmode
		} // end: Loop FT in Suite1
		suiteEndTime = TestUtil.now("dd.MMMMM.yyyy hh.mm.ss aaa");
	}
	
	
	/*public void TestCase()
	{
	for (int tcid = 2; tcid <= controller.getRowCount(currentFTID); tcid++)
		{
		// Check the TC Runmode
			if(controller.getCellData(currentFTID, "Runmode", tcid).equalsIgnoreCase("Y")) 
				{
					currentTCID = controller.getCellData(currentFTID, "TCID", tcid);
					currentTCName = controller.getCellData(currentFTID, "TC_Name", tcid);
					System.out.println("Executing Testcase ID = " + currentTCID);
			
					proceedOnFail=controller.getCellData(currentFTID, "ProceedOnFail", tcid);
					String ftPath = GLOBAL.getProperty("testcaseDir") + currentFTID + "\\";
					testStep = new Xlfile_Reader(System.getProperty("user.dir") + ftPath + currentTCID + ".xlsx");
			
					startTime=TestUtil.now("dd.MMMMM.yyyy hh.mm.ss aaa");
	
					// Create log file for each test case
					currentLogName = currentFTID + "_" + currentTCID + StartDate + GLOBAL.getProperty("logFileExt");
					TestReports.createExecutionLog(reportCurrentPath, currentLogName);
			
					// Initialize environment for the test case prior to execution
					DriverLib.testcaseEnter();

						for(int tsid = 2; tsid <= testStep.getRowCount(currentTCID); tsid++) 
							{
								String strTSID = testStep.getCellData(currentTCID, "TSID", tsid);
									if(strTSID != "") 
									{
										System.out.println("Executing Test Step = " + testStep.getCellData(currentTCID, "Decription", tsid));
										try
										{
											// Invoke a relevant method(keyword) to execute
											keyword=testStep.getCellData(currentTCID, "Keyword", tsid);
											object=testStep.getCellData(currentTCID, "Object", tsid);
											currentTSID=testStep.getCellData(currentTCID, "TSID", tsid);
											stepDescription=testStep.getCellData(currentTCID, "Decription", tsid);
											tsData=testStep.getCellData(currentTCID, "Data", tsid);
											if(tsData.contains(",")) {
												listTsData = Arrays.asList(tsData.split("\\s*,\\s*"));
											}
						
											Method method= KeywordsApp.class.getMethod(keyword);
											tsResult = (String)method.invoke(method);
						
										}
										catch(Throwable t)
										{
											System.out.println("Error came from " 
													+ "keyword= " + keyword
													+ " , object= " + object 
													+ " , testData= " + testData.getCellData(currentFTID, tsData, currentTCID));
											tsResult = FAIL;
											driver.quit();
						
										}
									}
								}
							}
				}
		}*/
				

	
	@AfterSuite
	public static void endScript() throws Exception{
		// Create the summary report
		String summaryreportFilePath = reportCurrentPath + "index" + GLOBAL.getProperty("reportFileExt");
		TestReports.createSummaryReport(summaryreportFilePath, featurelistFilePath);
		
		// Compress the latest report folder
		String compressedFileName = StartDate + ".zip";
		String compressedFilePath = reportPath + compressedFileName;
		TestUtil.zip(reportCurrentPath, reportPath + StartDate + ".zip");
		
		// Send the log report to the FTP server
		/*if(CONFIG.getProperty("log_will_upload").equalsIgnoreCase("y")) {
		TestReports.ftpUpload(CONFIG.getProperty("ftp_host"), CONFIG.getProperty("ftp_user"), 
				CONFIG.getProperty("ftp_password"), compressedFilePath, compressedFileName);
		}*/
		
		// Send email after test complete
		if(CONFIG.getProperty("mail_will_send").equalsIgnoreCase("y")) {
			String[] toAddresses = CONFIG.getProperty("mail_to").split(separator);
			String[] mail_attachFiles = {compressedFilePath};
			
			/*TestReports.sendMail(CONFIG.getProperty("mail_host"), CONFIG.getProperty("mail_port"), 
					CONFIG.getProperty("mail_from"), CONFIG.getProperty("mail_password"), 
					toAddresses, CONFIG.getProperty("mail_subject"), 
					CONFIG.getProperty("mail_message"), mail_attachFiles);*/
		}
	}
}

