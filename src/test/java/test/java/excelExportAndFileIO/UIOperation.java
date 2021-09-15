package test.java.excelExportAndFileIO;

import java.awt.Robot;
import java.awt.event.InputEvent;
import java.awt.event.KeyEvent;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;


//import org.openqa.selenium.security.UserAndPassword;
//import org.openqa.selenium.security.Credentials;



public class UIOperation {
	 static int urlRowNum=0;
	 static int listCellNum=1;
	 static int listCellNum2=1;
	 static int newTabCount=0;
	 static int setTextCount=0;
	 static int setTextRowCount=1;
	 static int setTextColumnCount=0;
	 static int currentColumn = 0;
	 static int trackFourthSheetRow=1;
	 static String temp="";
	 static String temp2="";	 
	 static String temp3="";
//	 static String status="FAIL";
	 
	 WebDriver driver;
	 String h1TagStatus="Something went wrong"; 
	 String metaTagStatus="Something went wrong";  
	 String pageTitleStatus="Something went wrong"; 
	 String getTextStatus="Something went wrong"; 
	 String redirectStatus="Something went wrong"; 
//	 String productTagStatus="Something went wrong";
	 String compareMultipleListStatus="";
//	 Actions action1=new Actions(driver);	
	 
	 WriteExcelFile w=new WriteExcelFile();
	 
	public UIOperation(WebDriver driver){
		this.driver = driver;
	}
	@SuppressWarnings({ "unchecked", "null" })
	public void perform(String testcaseName, String operation,String action, String locator, 
			String expression, String attribute, String value, String driverDelay, String expectedText, 
			String actualText, String status, String trackRow) throws Exception {
		try {
			status="FAIL";
			Sheet iQ=w.openFourthExcel();
			
			WebDriverWait wait=new WebDriverWait(driver, 20);
			System.out.println("");
			
			switch (operation.toUpperCase()) {
			case "CLICK":
				//Perform click
				driver.findElement(this.getObject(action,locator,expression)).click();		
				Thread.sleep(Long.parseLong(driverDelay));
				status="PASS";
			break;
			
			case "SETTEXT":
				//Set text on control
				if(!value.equals("NA"))
				{
					driver.findElement(this.getObject(action,locator,expression)).sendKeys(value);
					Thread.sleep(Long.parseLong(driverDelay));
				}
				else
				{
					int lastColumnInFourthSheet=iQ.getRow(0).getLastCellNum()-5;
					if(temp2.equals("browserIsClosed"))
					{
						if(setTextColumnCount<lastColumnInFourthSheet)
						{
							if(temp3.equals("setTextRowCountAlreadyIncremented"))
								setTextRowCount=setTextRowCount+0;
							else
								setTextRowCount=setTextRowCount+1;
							setTextColumnCount=0;
							temp="textColumnCountIsNotZero";
							System.out.println(temp);
							System.out.println("setTextRowCount= "+setTextRowCount);
							temp3="";
						}
//						temp="textColumnCountIsNotZero";
						temp2="browserIsNotClosed";
						System.out.println("IF-3");
					}
					
					driver.findElement(this.getObject(action,locator,expression)).clear();
					Thread.sleep(2000);
					iQ.getRow(setTextRowCount).getCell(setTextColumnCount).setCellType(CellType.STRING);
					String value1=iQ.getRow(setTextRowCount).getCell(setTextColumnCount).getStringCellValue();
					driver.findElement(this.getObject(action,locator,expression)).sendKeys(value1);
									
					System.out.println("Last column= "+iQ.getRow(0).getLastCellNum());
												
					if((setTextColumnCount==lastColumnInFourthSheet))
					{
						setTextRowCount=setTextRowCount+1;
						setTextColumnCount=0;	
						temp="textColumnCountZero";
						temp2="browserIsNotClosed";
						temp3="setTextRowCountAlreadyIncremented";
						System.out.println(temp);
						System.out.println("IF-1");
					}
					
					if((!temp.equals("textColumnCountZero")))
					{
						if(setTextColumnCount<lastColumnInFourthSheet)
						{
							setTextColumnCount=setTextColumnCount+1;
							System.out.println("Set Text Column Count= "+setTextColumnCount);
							temp="textColumnCountIsNotZero";
							temp3="";
							System.out.println("IF-2");
						}
					}
			   	}
				 Thread.sleep(Long.parseLong(driverDelay));
				 status="PASS";
				 break;
				
			case "GOTOURL":
				//Get url of application
				Sheet getUrlCountSheet=w.openZerothExcel();
				String getUrlCount=getUrlCountSheet.getRow(Integer.parseInt(trackRow)).getCell(2).getStringCellValue();
				
				if(getUrlCount.trim().equalsIgnoreCase("MULTIPLEURL"))
				{
					Sheet iQuantiSheet=w.openFirstExcel();
					urlRowNum=urlRowNum+1;
					listCellNum=1;
					listCellNum2=1;
					System.out.println("First sheet row count (1_URL)= "+urlRowNum);
					String url=iQuantiSheet.getRow(urlRowNum).getCell(0).getStringCellValue();
					driver.get(url);
				}
				else if(getUrlCount.equalsIgnoreCase("ONLYONEURL"))
				{
					urlRowNum=urlRowNum+1;
					listCellNum=1;
					listCellNum2=1;
					System.out.println("First sheet row count (1_URL)= "+urlRowNum);
					Sheet iQuantiSheet=w.openFirstExcel();
					String url=iQuantiSheet.getRow(1).getCell(0).getStringCellValue();
					driver.get(url);
				}
				else if(getUrlCount.equalsIgnoreCase("SPECIFYURLHERE"))
				{
					urlRowNum=urlRowNum+1;
					listCellNum=1;
					listCellNum2=1;
					System.out.println("First sheet row count (1_URL)= "+urlRowNum);
					driver.get(value);
				}
								
//				UsernamePasswordCredentials UP = new UsernamePasswordCredentials("punit","punit");
//				Alert alert=driver.switchTo().alert();
//				alert.authenticateUsing(UP);
				
//				Alert alert2=wait.until(ExpectedConditions.alertIsPresent());  
//				alert2.authenticateUsing(new UserAndPassword("cheese",  "secretGouda"));
				
//				UsernamePasswordCredentials UP = new UsernamePasswordCredentials("punit","punit");
//				driver.switchTo().alert().authenticateUsing(UP);
				
//				Runtime.getRuntime().exec("E:\\Java_Programs\\HybridFramework\\Login Credentials\\loginCredentialsCHIP.exe");
				
				driver.switchTo().defaultContent();
				
//				new WebDriverWait(driver, 30).until(
//					      webDriver -> ((JavascriptExecutor) webDriver).executeScript("return document.readyState").equals("complete"));
				
				Thread.sleep(Long.parseLong(driverDelay));
				status="PASS";
				break;
			
			case "GOBACK":
					driver.navigate().back();
					Thread.sleep(Integer.parseInt(driverDelay));
					status="PASS";
				break;
				
			case "GOFORWARD":
					driver.navigate().forward();
					Thread.sleep(Long.parseLong(driverDelay));
					status="PASS";
					break;
			
			case "OPENNEWTAB":
				newTabCount=newTabCount+1; 
				Robot robot1 = new Robot();                          
				robot1.keyPress(KeyEvent.VK_CONTROL); 
				robot1.keyPress(KeyEvent.VK_T); 
				robot1.keyRelease(KeyEvent.VK_CONTROL); 
				robot1.keyRelease(KeyEvent.VK_T);

				Thread.sleep(2000);
				//Switch focus to new tab
				ArrayList<String> tabs1 = new ArrayList<String> (driver.getWindowHandles());
				driver.switchTo().window(tabs1.get(newTabCount));
				
				Thread.sleep(Long.parseLong(driverDelay));
				status="PASS";
				break;
				
			case "JUMPTOTAB":
//				Actions action1=new Actions(driver);
//				action1.keyDown(Keys.CONTROL).sendKeys(Keys.TAB).build().perform();
				
//				driver.switchTo().window(Keys.chord(Keys.CONTROL,"\t"));
				
//				driver.switchTo().window(Keys.chord(Keys.CONTROL,"\t"));
				
//				Thread.sleep(3000);
//				driver.switchTo().window(Keys.chord(Keys.CONTROL,"\t"));

//				Robot robot2 = new Robot();                          
//				robot2.keyPress(KeyEvent.VK_CONTROL); 
//				robot2.keyPress(KeyEvent.VK_TAB); 
//				robot2.keyRelease(KeyEvent.VK_CONTROL); 
//				robot2.keyRelease(KeyEvent.VK_TAB);
				
				ArrayList<String> tabs2 = new ArrayList<String> (driver.getWindowHandles());
				driver.switchTo().window(tabs2.get(Integer.parseInt(value)));
				
				Thread.sleep(Long.parseLong(driverDelay));
				status="PASS";
				break;
				
//			case "SWITCHTABRTL":
//				Actions action2=new Actions(driver);
//				action2.keyDown(Keys.CONTROL).sendKeys(Keys.SHIFT).sendKeys(Keys.TAB).build().perform();
	//
//				status="PASS";
//				w.openZerothExcel();
//				w.writeZerothExcelStatus(Integer.parseInt(trackRow), status);
//				w.closeExcel();
//				break;
				
			case "WAITUNTILPAGELOADS":
//				List<WebElement> waitUntil=driver.findElements(this.getObject(action, locator, expression);
				
				driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
//				new WebDriverWait(driver, 5).until(
//					      webDriver -> ((JavascriptExecutor) webDriver).executeScript("return document.readyState").equals("complete"));
					
				Thread.sleep(Long.parseLong(driverDelay));
				status="PASS";
				break;
				
			case "WAITUNTILELEMENTISVISIBLE":
				WebElement web=driver.findElement(this.getObject(action, locator, expression));
				
				wait.until(ExpectedConditions.visibilityOf(web));
				Thread.sleep(Long.parseLong(driverDelay));
				status="PASS";
				break;
				
			case "GETTEXT":
				//Get text of an element
				actualText=driver.findElement(this.getObject(action,locator, expression)).getText();
				System.out.println("Actual Text= "+actualText);
				Thread.sleep(Long.parseLong(driverDelay));
				
				if(value.equalsIgnoreCase("NA"))
				{
					int actualGetTextColumnCount=iQ.getRow(0).getLastCellNum()-3;
					int expectedTextColumn=iQ.getRow(0).getLastCellNum()-4;
					System.out.println("Expected Text Column= "+expectedTextColumn);
					
					String expectedText2=w.openFourthExcel().getRow(trackFourthSheetRow).getCell(expectedTextColumn).getStringCellValue();
					System.out.println("Get Text from Fourth Sheet: "+expectedText2);
					
					if(!expectedText2.trim().equalsIgnoreCase("NA"))
					{
						if(expectedText2.trim().equals(actualText))
						{
							status="PASS";
						}
						else
						{
							status="FAIL";
						}
					}
					else
					{
						status="Value Fetched!";
					}
					w.openFourthExcel();
					w.writeActualGetText2(trackFourthSheetRow, actualGetTextColumnCount, actualText, status);
					w.closeExcel();
					trackFourthSheetRow=trackFourthSheetRow+1;	
				}
				else
				{
								
					expectedText=w.openZerothExcel().getRow(Integer.parseInt(trackRow)).getCell(8).getStringCellValue();
					if(expectedText.trim().equals(actualText))
					{
						 status="PASS";
					}
					else
					{
						 status="FAIL";
					}
					w.openZerothExcel();
					w.writeActualGetText(Integer.parseInt(trackRow), actualText);
					w.closeExcel();
				}
				Thread.sleep(Long.parseLong(driverDelay));
				 status="PASS";
				break;
				
			case "DROPDOWN":
				//Set text on control
				List<WebElement> dropDown=driver.findElements(this.getObject(action,locator,expression));
				String selectedOption="Something went wrong";
				selectedOption=dropDown.get(Integer.parseInt(value)).getText().toString();	
				dropDown.get(Integer.parseInt(value)).click();
				
				Thread.sleep(Long.parseLong(driverDelay));
				status="PASS";
				break;
				
			case "ONEFROMLIST":
				//Set text on control
				List<WebElement> list=driver.findElements(this.getObject(action,locator,expression));
				String listItem="Something went wrong";
				
				listItem=list.get(Integer.parseInt(value)).getAttribute(attribute.toString());	
//					list.get(Integer.parseInt(value)).click();
				
				Thread.sleep(Long.parseLong(driverDelay));
				status="PASS";
				break;
			
			case "LISTITEMCLICK":
				//Set text on control
				List<WebElement> list2=driver.findElements(this.getObject(action,locator,expression));
				String listItem2="Something went wrong";
				//listItem2=list2.get(Integer.parseInt(value)).getAttribute(attribute.toString());	
				list2.get(Integer.parseInt(value)).click();
				
				Thread.sleep(Long.parseLong(driverDelay));
				status="PASS";
				break;
				
			case "H1":
				WebElement wb=driver.findElement(this.getObject(action, locator, expression));
				String actualH1Tag=wb.getText();
				System.out.println("H1 Tag Name is "+actualH1Tag);
				String expectedH1Tag=w.openFirstExcel().getRow(urlRowNum).getCell(1).getStringCellValue();
				if(expectedH1Tag.trim().equals(actualH1Tag))
					h1TagStatus="PASS";
				else
					h1TagStatus="FAIL";
				
				w.openFirstExcel();
				w.writeH1Tag(urlRowNum, actualH1Tag, h1TagStatus);
				w.closeExcel();
				
				Thread.sleep(Long.parseLong(driverDelay));
				status="PASS";
				break;
			
			case "META":
				wb=driver.findElement(this.getObject(action, locator, expression));
				String actualMetaTag=wb.getAttribute("content");
				System.out.println("H1 Tag Name is "+actualMetaTag);
				String expectedMetaTag=w.openFirstExcel().getRow(urlRowNum).getCell(4).getStringCellValue();
				if(expectedMetaTag.trim().equals(actualMetaTag))
					metaTagStatus="PASS";
				else
					metaTagStatus="FAIL";
				
				w.openFirstExcel();
				w.writeMetaTag(urlRowNum, actualMetaTag, metaTagStatus);
				w.closeExcel();
				
				Thread.sleep(Long.parseLong(driverDelay));
				status="PASS";
				break;
			
			case "PAGETITLE":
				String actualPageTitle=driver.getTitle();
				System.out.println("H1 Tag Name is "+actualPageTitle);
				String expectedPageTitle=w.openFirstExcel().getRow(urlRowNum).getCell(7).getStringCellValue();
				if(expectedPageTitle.trim().equals(actualPageTitle))
					pageTitleStatus="PASS";
				else
					pageTitleStatus="FAIL";
				
				w.openFirstExcel();
				w.writePageTitle(urlRowNum, actualPageTitle, pageTitleStatus);
				w.closeExcel();
				
				Thread.sleep(Long.parseLong(driverDelay));
				status="PASS";
				break;
			
			case "REDIRECT":
				String actualUrl=driver.getCurrentUrl();
				System.out.println("Actual is "+actualUrl);
				String expectedUrl=w.openFirstExcel().getRow(urlRowNum).getCell(10).getStringCellValue();
				if(expectedUrl.trim().equals(actualUrl))
					redirectStatus="PASS";
				else
					redirectStatus="FAIL";
				w.openFirstExcel();
				w.writeRedirectedUrl(urlRowNum, actualUrl, redirectStatus);
				w.closeExcel();
				
				Thread.sleep(Long.parseLong(driverDelay));
				status="PASS";
				break;
				
			case "COMPARETMANYTOMANY":
				List<WebElement> productTags=driver.findElements(this.getObject(action, locator, expression));
				List<String> allProductTags=new ArrayList<String>();
				String actualProductTags="";
				listCellNum2=listCellNum2+0;
				String productTagStatus="Something went wrong";
				
				String expectedProductTags=w.openThirdExcel().getRow(urlRowNum).getCell(listCellNum2).getStringCellValue();
				
				for(int i=0;i<productTags.size();i++)
					allProductTags.add(productTags.get(i).getText());
				
				for(int j=0;j<allProductTags.size();j++)
					actualProductTags=actualProductTags + allProductTags.get(j)+"^ ";
				
//				Backslash (\) issue is resolved but uppercase & lowercase issue is yet to be resolved 
				int expectedProductTagsLength=expectedProductTags.trim().replaceAll("[^a-zA-Z0-9\\\\\\\\\\.:\\'_()\\-&@$%/\\*]", "").length();
				System.out.println("Expected Product Tag Replaced= "+ expectedProductTags.trim().replaceAll("[^a-zA-Z0-9\\\\\\\\\\.:\\'_()\\-&@$%/\\*]", ""));
				int actualProductTagsLength=actualProductTags.replaceAll("[^a-zA-Z0-9\\\\\\\\\\.:\\'_()\\-&@$%/\\*]", "").length();
				
				System.out.println("expectedProductTagsLength= "+expectedProductTagsLength);
				System.out.println("actualProductTagsLength= "+actualProductTagsLength);
				
				if(!expectedProductTags.equalsIgnoreCase("NA"))
				{
					for(int j=0;j<productTags.size();j++)
					{
						if(actualProductTagsLength==expectedProductTagsLength)
						{
							if(expectedProductTags.trim().contains(allProductTags.get(j)))
							{
								productTagStatus = "PASS";
							}
						}
						else
						{
							productTagStatus="FAIL";
							break;
						}
					}
				}
				else
				{
					productTagStatus="TEXT FETCHED";
				}
				
				w.openThirdExcel();
				w.writeProductTags(urlRowNum, listCellNum2, actualProductTags, productTagStatus);
				w.closeExcel();
				
				listCellNum2=listCellNum2+3;
				productTags.clear();
				allProductTags.clear();
				
				Thread.sleep(Long.parseLong(driverDelay));
				status="PASS";
				break;
				
			case "COMPAREAONETOMANY":
				List<WebElement> getAttributes=driver.findElements(this.getObject(action, locator, expression));
				List<String> actualGetLinks=new ArrayList<String>();
				String allGetLinks = "";
				listCellNum=listCellNum+0;
				System.out.println(listCellNum);
				
				for(int i=0;i<getAttributes.size();i++)
					actualGetLinks.add(getAttributes.get(i).getAttribute(attribute)) ;
				
				for(int i=0;i<getAttributes.size();i++)
					allGetLinks=allGetLinks + actualGetLinks.get(i) +", ";
				
				String expectedGetLinks=w.openSecondExcel().getRow(urlRowNum).getCell(listCellNum).getStringCellValue();
				System.out.println(expectedGetLinks);
				if(!expectedGetLinks.trim().equalsIgnoreCase("NA"))
				{
					for(int j=0;j<getAttributes.size();j++)
					{
						if(expectedGetLinks.trim().equals(actualGetLinks.get(j)))
						{
							compareMultipleListStatus = "PASS";
						}
						else
						{
							compareMultipleListStatus="FAIL";
							break;
						}
					}
				}
				else
				{
					compareMultipleListStatus = "Attribute's Value Fetched";
				}
				w.openSecondExcel();
				w.writeCompareMultipleList(urlRowNum, listCellNum, allGetLinks, compareMultipleListStatus);
				w.closeExcel();
				
				listCellNum=listCellNum+3;
				getAttributes.clear();
				actualGetLinks.clear();
				//Write Status to Zeroth sheet
				Thread.sleep(Long.parseLong(driverDelay));
				status="PASS";
				 break;
				 
			case "EXPLICITDELAY":
				Thread.sleep(Long.parseLong(driverDelay));
				status="PASS";
				break;
			
			case "ALERTOK":
				driver.switchTo().alert().accept();
				driver.switchTo().window(driver.getWindowHandle());
				Thread.sleep(Long.parseLong(driverDelay));
				status="PASS";
				break;
				
			case "ALERTCANCEL":
				driver.switchTo().alert().dismiss();
				Thread.sleep(Long.parseLong(driverDelay));
				status="PASS";
				break;
				
			case "MOUSETEXT":
				Robot robotText=new Robot();
//				robot.mouseMove(633, 183);
				robotText.mouseMove(Integer.parseInt(expectedText), Integer.parseInt(actualText));
				robotText.mousePress((int) InputEvent.TEXT_EVENT_MASK);
				Thread.sleep(1000);
				robotText.mouseRelease((int) InputEvent.TEXT_EVENT_MASK);
				Thread.sleep(1000);
				Thread.sleep(Long.parseLong(driverDelay));
				status="PASS";
				break;
				
			case "MOUSEBUTTON":
				Robot robotButton=new Robot();
//				robotButton.mouseMove(850, 283);
				robotButton.mouseMove(Integer.parseInt(expectedText), Integer.parseInt(actualText));
				robotButton.mousePress(InputEvent.BUTTON1_DOWN_MASK);
				Thread.sleep(1000);
				robotButton.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
				Thread.sleep(1000);
				
				Thread.sleep(Long.parseLong(driverDelay));
				status="PASS";
				break;
				
			case "USERNAME":
				Robot robotUserName=new Robot();
				Sheet fifthSheetUsername=w.openFifthExcel();
				 String[] username = new String[10];
				 int[] keyEventUsername = new int[10];
				 
//				username=fifthSheet.getRow(1).getCell(1).getStringCellValue();
//				password=fifthSheet.getRow(1).getCell(2).getStringCellValue();
				
//				System.out.println(password);
				int fifthSheetUsernameLastColumnNum=fifthSheetUsername.getRow(1).getLastCellNum();
				System.out.println(fifthSheetUsernameLastColumnNum);
//				
				for(int i=0;i<fifthSheetUsernameLastColumnNum-1;i++) {
					username[i]=fifthSheetUsername.getRow(1).getCell(i+1).getStringCellValue();
				}
				
				for(int i=0;i<fifthSheetUsernameLastColumnNum-1;i++) {
					System.out.println(username[i]);
//					System.out.println(Integer.valueOf(username[i],32));
				}
				
				for(int i=0;i<fifthSheetUsernameLastColumnNum-1;i++) {
					Field f= KeyEvent.class.getField(username[i]);
					keyEventUsername[i]=f.getInt(null);
					System.out.println(keyEventUsername[i]);
//					System.out.println(Integer.valueOf(username[i],32));
				}
				
				for(int i=0;i<fifthSheetUsernameLastColumnNum-1;i++) {
					robotUserName.keyPress(keyEventUsername[i]);
				}
				
//				int[] key1= {Integer.parseInt(username)};
//				for (int i=0;i<key1.length;i++)
//				{
//					robotUserName.keyPress(username[i]);
//					Thread.sleep(1000);
//				}	

				Thread.sleep(Long.parseLong(driverDelay));
				status="PASS";				
				break;
				
			case "PASSWORD":
				Robot robotPassword=new Robot();
				Sheet fifthSheetPassword=w.openFifthExcel();
				String[] password = new String[10];
				int[] keyEventPassword = new int[10];
				
//				int[] key2= {KeyEvent.VK_P, KeyEvent.VK_U,KeyEvent.VK_N,KeyEvent.VK_I,KeyEvent.VK_T};
//				for (int i=0;i<key2.length;i++)
//				{
//					robotPassword.keyPress(key2[i]);
////					Thread.sleep(1000);
//				}		
				
				int fifthSheetPasswordLastColumnNum=fifthSheetPassword.getRow(2).getLastCellNum();
				System.out.println(fifthSheetPasswordLastColumnNum);
//				
				for(int i=0;i<fifthSheetPasswordLastColumnNum-1;i++) {
					password[i]=fifthSheetPassword.getRow(2).getCell(i+1).getStringCellValue();
				}
				
				for(int i=0;i<fifthSheetPasswordLastColumnNum-1;i++) {
					System.out.println(password[i]);
				}
				
				for(int i=0;i<fifthSheetPasswordLastColumnNum-1;i++) {
					Field f= KeyEvent.class.getField(password[i]);
					keyEventPassword[i]=f.getInt(null);
					System.out.println(keyEventPassword[i]);
				}
				
				for(int i=0;i<fifthSheetPasswordLastColumnNum-1;i++) {
					robotPassword.keyPress(keyEventPassword[i]);
				}

				 Thread.sleep(Long.parseLong(driverDelay));
				 status="PASS";	
				break;
				
			case "ENTERKEY":
				Robot robotEnter=new Robot();
				robotEnter.keyPress(KeyEvent.VK_ENTER);
				Thread.sleep(1000);
				robotEnter.keyRelease(KeyEvent.VK_ENTER);
				Thread.sleep(1000);

				Thread.sleep(Long.parseLong(driverDelay));
				status="PASS";				 
				break;	
				
			case "TAB":
				Robot robotTAB=new Robot();
				robotTAB.keyPress(KeyEvent.VK_TAB);
				Thread.sleep(1000);
				robotTAB.keyRelease(KeyEvent.VK_TAB);
				Thread.sleep(1000);
				Thread.sleep(Long.parseLong(driverDelay));
				status="PASS";	
				break;
				
			case "CLOSE":
				driver.quit();
				temp2="browserIsClosed";
				Thread.sleep(Long.parseLong(driverDelay));
				 status="PASS";
				 break;
				 
			default:
				if(operation.equalsIgnoreCase("CHROME")||operation.equalsIgnoreCase("FIREFOX")||operation.equalsIgnoreCase("EDGE"))
					status="PASS";
				else
					status="FAIL";
				break;
			}
		}
		finally {
//			System.out.println("finally");
			w.openZerothExcel();
			 w.writeZerothExcelStatus(Integer.parseInt(trackRow), status);
			 w.closeExcel();
		}
		
	}
	
	/**
	 * Find element BY using object type and value
	 * @param p
	 * @param objectName
	 * @param objectType
	 * @return
	 * @throws Exception
	 */
	private By getObject(String action, String locator, String expression) throws Exception {
		//Find by xpath
		if(locator.equalsIgnoreCase("XPATH")){
			
			return By.xpath(expression);
		}
		
		//find by ID
		else if(locator.equalsIgnoreCase("ID")){
			return By.id(expression);
		}
		
		//find by class
		else if(locator.equalsIgnoreCase("CLASSNAME")){
			
//			return By.className(p.getProperty(objectName));
			return By.className(expression);
		}
		//find by name
		else if(locator.equalsIgnoreCase("NAME")){
			
//			return By.name(p.getProperty(objectName));
			return By.name(expression);	
		}
		//Find by css
		else if(locator.equalsIgnoreCase("CSS")){
			
//			return By.cssSelector(p.getProperty(objectName));
			return By.cssSelector(expression);
		}
		//find by link
		else if(locator.equalsIgnoreCase("LINK")){
			
//			return By.linkText(p.getProperty(objectName));
			return By.linkText(expression);
		}
		//find by partial link
		else if(locator.equalsIgnoreCase("PARTIALLINK")){
//			return By.partialLinkText(p.getProperty(objectName));
			return By.partialLinkText(expression);
		}
		else
		{
			throw new Exception("Wrong object type");
		}
	}
}
