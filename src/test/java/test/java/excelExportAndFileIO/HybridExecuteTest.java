package test.java.excelExportAndFileIO;

import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.testng.TestListenerAdapter;
import org.testng.TestNG;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;



public class HybridExecuteTest {
	WebDriver webdriver = null;
	
	static int invocationCount=1;

	public static int rowCount=0;
	public static int rowCount2=0;
	public static int totalRows=0;
	static String browser;
	static String driverPath;
	
	static int zerothSheetLastRowNum=0;
	static int fourthSheetLastRowNum=0;
	

	//This is the main method if you want to create .EXE file of this project and this is optional
	public static void main(String[] args) {
		TestListenerAdapter tla = new TestListenerAdapter();
		TestNG testng = new TestNG();
		testng.setTestClasses(new Class[] { HybridExecuteTest.class });
		testng.addListener(tla);
		testng.run();
		}
	
	@BeforeTest
	public void beforeTest() throws IOException, InterruptedException
	{
		WriteExcelFile w=new WriteExcelFile();
		w.openMasterExcel();
	    Sheet excel=w.openZerothExcel();
	    Sheet fourthSheet=w.openFourthExcel();
	    zerothSheetLastRowNum=excel.getLastRowNum();
	    System.out.println("zerothSheetLastRowNum= "+zerothSheetLastRowNum);
	    
    	browser=excel.getRow(1).getCell(1).getStringCellValue();
    	driverPath=excel.getRow(1).getCell(2).getStringCellValue();
    	System.out.println("Driver path= "+driverPath);
    	zerothSheetLastRowNum=excel.getLastRowNum();
    	fourthSheetLastRowNum=fourthSheet.getLastRowNum();
    	
//    	for(int i=1;i<zerothSheetLastRowNum+1;i++)
//    	{
//    		w.openZerothExcel();
//    		w.writeZerothExcelStatus(i, "NA");
//    		w.closeExcel();
//    	}
    	Thread.sleep(2000);
    	System.out.println("Before Test Ends");
	}
	
    @Test(dataProvider="hybridData", invocationCount = 1, testName = "SEO and Functional Test" )
	public void testBegins(String testcaseName,String keyword,String action,String locator,
			String expression, String attribute, String value, String driverDelay, String expectedText, 
			String actualText, String status, String trackRow) throws Exception {
		// TODO Auto-generated method stub
		rowCount=rowCount+1;
		rowCount2=rowCount2+1;
    	if(rowCount2==zerothSheetLastRowNum-1)
//    	{
//    		invocationCount=invocationCount+1;
//    		rowCount2=0;
//    	}
    	System.out.println("Row Count= "+rowCount+" out of "+totalRows*invocationCount);

    	if(testcaseName!=null&&testcaseName.length()!=0){
    		if(!driverPath.trim().equalsIgnoreCase("NA")) {
    			status="FAIL";
		    	if(browser.equalsIgnoreCase("chrome"))
				{
					System.setProperty("webdriver.chrome.driver", driverPath);
					webdriver=new ChromeDriver();
					status="PASS";
				}
				else if(browser.equalsIgnoreCase("firefox"))
				{
					System.setProperty("webdriver.gecko.driver", driverPath);
					webdriver=new FirefoxDriver();
					status="PASS";
				}
				else if(browser.equalsIgnoreCase("edge"))
				{
					System.setProperty("webdriver.edge.driver", driverPath);
					webdriver=new EdgeDriver();
					status="PASS";
				}
				else if(browser.equalsIgnoreCase("ie"))
				{
					System.setProperty("webdriver.ie.driver", driverPath);
					webdriver=new InternetExplorerDriver();
					status="PASS";
				}
	    	}
	    	else
	    	{
//    	    		System.out.println("User dir= "+System.getProperty("user.dir"));
	    		status="FAIL";
	    		if(browser.equalsIgnoreCase("chrome"))
				{
					String path=System.getProperty("user.dir");
//    					System.out.println("Chrome driver path= "+ path);
					System.setProperty("webdriver.chrome.driver", path+"\\drivers\\chromedriver.exe");
					webdriver=new ChromeDriver();
					status="PASS";
				}
				else if(browser.equalsIgnoreCase("firefox"))
				{
	//				System.getProperty("user.dir"+"\\", "chromedriver.exe");
					String firefoxDriver=System.getProperty("user.dir");
					System.setProperty("webdriver.firefox.driver", firefoxDriver+"\\drivers\\geckodriver.exe");
					webdriver=new FirefoxDriver();
					status="PASS";
				}
				else if(browser.equalsIgnoreCase("edge"))
				{
					String edgeDriver=System.getProperty("user.dir");
					System.setProperty("webdriver.edge.driver", edgeDriver+ "\\drivers\\msedgedriver.exe");
					webdriver=new EdgeDriver();
					status="PASS";
				}
				else if(browser.equalsIgnoreCase("ie"))
				{
					String ieDriver=System.getProperty("user.dir");
					System.setProperty("webdriver.ie.driver", ieDriver+ "\\drivers\\IEDriverServer.exe");
					webdriver=new InternetExplorerDriver();
					status="PASS";
    				}
    	    	}
        		webdriver.manage().window().maximize();
            	webdriver.manage().deleteAllCookies();
            	webdriver.manage().timeouts().pageLoadTimeout(40, TimeUnit.SECONDS);
            	webdriver.manage().timeouts().implicitlyWait(40, TimeUnit.SECONDS);
        	}
   
            UIOperation operation = new UIOperation(webdriver);
          	//Call perform function to perform operation on UI
        
    	operation.perform(testcaseName, keyword, action, locator, 
    			expression, attribute, value, driverDelay, expectedText, actualText, 
    			status,trackRow);    	
	}
    
    @SuppressWarnings("deprecation")
	@DataProvider(name="hybridData")
	public Object[][] getDataFromDataprovider() throws IOException {
    	Object[][] object = null; 
//    	ReadExcelFile file = new ReadExcelFile();
        
         //Read keyword sheet
//         Sheet iQuantiSheet = file.readExcel(System.getProperty("user.dir")+"\\","TestCase.xlsx", "0");
    	WriteExcelFile w=new WriteExcelFile();
    	w.openMasterExcel();
    	Sheet iQuantiSheet=w.openZerothExcel();
       //Find number of rows in excel file
     	int rowCount = iQuantiSheet.getLastRowNum()-iQuantiSheet.getFirstRowNum();
     	object = new Object[rowCount][iQuantiSheet.getRow(0).getLastCellNum()];  //OR object = new Object[rowCount][7];
     	for (int i = 0; i < rowCount; i++) {
    		//Loop over all the rows
    		Row row = iQuantiSheet.getRow(i+1);
    		//Create a loop to print cell values in a row
    		for (int j = 0; j < row.getLastCellNum(); j++) {
    			//Print excel data in console
    			row.getCell(j).setCellType(CellType.STRING);
    			object[i][j] = row.getCell(j).toString();
    		}
    	}
//     	System.out.println("");
     	System.out.println(object.length);
     	totalRows=object.length;
     	return object;	 
	}
}
