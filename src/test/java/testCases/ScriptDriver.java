package testCases;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.concurrent.TimeUnit;
import java.util.stream.Collectors;

import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.Point;
import org.openqa.selenium.TakesScreenshot;
import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

import Pages.searchpage;

public class ScriptDriver {
	WebDriver driver;
	ExtentTest test;
	ExtentReports report;
	

	@Test(dataProvider = "ToolData",threadPoolSize=2)
	public void driverSetup(HashMap<String, String> data) throws InterruptedException, IOException {
		// TODO Auto-generated method stub
		try {
			
		report = new ExtentReports("//Users//ramkumars//eclipse-workspace//Selenium_Framework2//src//Results//Amazonpurchaseresults.html");
		test = report.startTest("ExtentDemo");
		
		System.out.print("Got the data value :"+data.get("ToRun"));
		//Firefox driver
		/* 
		WebDriver driver = new FirefoxDriver();
		driver.manage().window().maximize();
		driver.manage().deleteAllCookies();
		driver.manage().timeouts().pageLoadTimeout(60, TimeUnit.SECONDS);
		driver.get("http://google.com");
		driver.quit();*/
		
		
		//Chrome driver
		System.setProperty("webdriver.chrome.driver", "/Users/ramkumars/eclipse-workspace/UIComparision/src/test/java/Alldrivers/chromedriver");
		
		
		ChromeOptions options = new ChromeOptions();
		options.addArguments("--remote-allow-origins=*");
		driver = new ChromeDriver(options);
		
		//driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.manage().timeouts().pageLoadTimeout(60, TimeUnit.SECONDS);
		driver.get(data.get("URL"));
		
		//To get all Xpath links
		getxpathjscript("//a");
		//getxpathjscript("//input");
		//getxpathjscript("//*");
		
		driver.get("file:////"+data.get("URL2"));
		comparepages();
		
		report.endTest(test);
		report.flush();
		this.takeSnapShot(driver, "/Users/ramkumars/eclipse-workspace/UIComparision/src/test/java/Results/check"+appenddatetime()+".png") ;
		driver.quit();
		}catch(Exception e) {
			e.printStackTrace();
		}
		
	
	}
	
	//---------------------------------------------------------------------------------
	//---------------------------------------------------------------------------------
	
	// Javascript to collect all links
	public void getxpathjscript(String collectxpath) throws IOException {
				java.util.List<WebElement> links = driver.findElements(By.xpath(collectxpath));
			      System.out.println("Number of Elements in the Page is " + links.size());
			      
			      
			      File file =    new File("/Users/ramkumars/eclipse-workspace/UIComparision/src/test/java/TestData/ToolsQATestData3.xls");
			        
			        //Create an object of FileInputStream class to read excel file
			        FileInputStream inputStream = new FileInputStream(file);
			        
			        //creating workbook instance that refers to .xls file
			        HSSFWorkbook wb=new HSSFWorkbook(inputStream);
			        
			        //creating a Sheet object using the sheet Name
			        HSSFSheet sheet=wb.getSheet("ActualPage");
			        			      
			      
			      for (int i = 1; i<=links.size()-1; i = i+1) {
			    	  
			    	//Create a row object to retrieve row at index 3
				        HSSFRow row2=sheet.createRow(i);
				   
			    	  
			         String xpath = AbsoluteXPath(links.get(i), driver);
					 System.out.println(xpath);
					 String Textvalue = driver.findElement(By.xpath(xpath)).getText();
					 System.out.println("Text Value : "+Textvalue);
					 String fontsize = driver.findElement(By.xpath(xpath)).getCssValue("font-size");
					 System.out.println("Font Size : "+fontsize);
					 String fontcolor = driver.findElement(By.xpath(xpath)).getCssValue("color");
					 System.out.println("Font Color : "+fontcolor);
					 String backgroundcolor = driver.findElement(By.xpath(xpath)).getCssValue("background-color");
					 System.out.println("Background Color : "+backgroundcolor);
					 String fontfamily = driver.findElement(By.xpath(xpath)).getCssValue("font-family");
					 System.out.println("Font Family : "+fontfamily);
					 Point point = driver.findElement(By.xpath(xpath)).getLocation();
				        System.out.println("x coordinate: " + point.getX());
				        System.out.println("y coordinate: " + point.getY());
				        
				        //create a cell object to enter value in it using cell Index
				        row2.createCell(0).setCellValue(i);
				        row2.createCell(1).setCellValue(xpath);
				        row2.createCell(2).setCellValue(Textvalue);
				        row2.createCell(3).setCellValue(fontsize);
				        row2.createCell(4).setCellValue(fontcolor);
				        row2.createCell(5).setCellValue(backgroundcolor);
				        row2.createCell(6).setCellValue(fontfamily);
				        row2.createCell(7).setCellValue(point.getX());
				        row2.createCell(8).setCellValue(point.getY());
			      }	
			        //write the data in excel using output stream
			        FileOutputStream outputStream = new FileOutputStream("/Users/ramkumars/eclipse-workspace/UIComparision/src/test/java/TestData/ToolsQATestData3.xls");
			        wb.write(outputStream);
			        outputStream.close();
			    
	}

	@SuppressWarnings("unused")
	public void comparepages() throws IOException {
		File file =    new File("/Users/ramkumars/eclipse-workspace/UIComparision/src/test/java/TestData/ToolsQATestData3.xls");
        
        //Create an object of FileInputStream class to read excel file
        FileInputStream inputStream = new FileInputStream(file);
        
        //creating workbook instance that refers to .xls file
        HSSFWorkbook wb=new HSSFWorkbook(inputStream);
        
        //creating a Sheet object using the sheet Name
        HSSFSheet sheet=wb.getSheet("ActualPage");
        
        
        int rowCount = sheet.getLastRowNum();
        System.out.println(rowCount);

        for(int i=0;i<=(rowCount-1);i++){
        
	        Row row = sheet.getRow(i+1);
	        Cell cell = row.getCell(1);
	        String cellval = cell.getStringCellValue();
			System.out.println(cellval);
			boolean present;
			try {
				driver.findElement(By.xpath(cellval));
			    present = true;
			    System.out.println(present);
			    
			     String ATextvalue = driver.findElement(By.xpath(cellval)).getText();
				 System.out.println("Actual Text Value : "+ATextvalue);
				 String ETextvalue = sheet.getRow(i+1).getCell(2).getStringCellValue();
				 System.out.println("Expected Text Value : "+ETextvalue);
				 if (ATextvalue.equals(ETextvalue)){
					 row.createCell(17).setCellValue("Matches");
				 }else {
					 row.createCell(17).setCellValue("Not Matches");
					 highlight(cellval);
				 }
				 
				 String Afontsize = driver.findElement(By.xpath(cellval)).getCssValue("font-size");
				 System.out.println("Actual Font Size : "+Afontsize);
				 String Efontsize = sheet.getRow(i+1).getCell(3).getStringCellValue();
				 System.out.println("Expected Font Size : "+Efontsize);
				 if (Afontsize.equals(Efontsize)){
					 row.createCell(18).setCellValue("Matches");
				 }else {
					 row.createCell(18).setCellValue("Not Matches");
					 highlight(cellval);
				 }
				 
				 
				 String Afontcolor = driver.findElement(By.xpath(cellval)).getCssValue("color");
				 System.out.println("Actual Font Color : "+Afontcolor);
				 String Efontcolor = sheet.getRow(i+1).getCell(4).getStringCellValue();
				 System.out.println("Expected Font Color : "+Efontcolor);
				 if (Afontcolor.equals(Efontcolor)){
					 row.createCell(19).setCellValue("Matches");
				 }else {
					 row.createCell(19).setCellValue("Not Matches");
					 highlight(cellval);
				 }
				 
				 String Abackgroundcolor = driver.findElement(By.xpath(cellval)).getCssValue("background-color");
				 System.out.println("Actual Background Color : "+Abackgroundcolor);
				 String Ebackgroundcolor = sheet.getRow(i+1).getCell(5).getStringCellValue();
				 System.out.println("Expected Font Color : "+Ebackgroundcolor);
				 if (Abackgroundcolor.equals(Ebackgroundcolor)){
					 row.createCell(20).setCellValue("Matches");
				 }else {
					 row.createCell(20).setCellValue("Not Matches");
					 highlight(cellval);
				 }
				 
				 
				 String Afontfamily = driver.findElement(By.xpath(cellval)).getCssValue("font-family");
				 System.out.println("Actual Font Family : "+Afontfamily);
				 String Efontfamily = sheet.getRow(i+1).getCell(6).getStringCellValue();
				 System.out.println("Expected Font Color : "+Efontfamily);
				 if (Afontfamily.equals(Efontfamily)){
					 row.createCell(21).setCellValue("Matches");
				 }else {
					 row.createCell(21).setCellValue("Not Matches");
					 highlight(cellval);
					 }
				 
				 Point Apoint = driver.findElement(By.xpath(cellval)).getLocation();
			        System.out.println("Actual x coordinate: " + Apoint.getX());
			        System.out.println("Actual y coordinate: " + Apoint.getY());
			        String Axcoordinate = String.valueOf(Apoint.getX());
			        String Aycoordinate = String.valueOf(Apoint.getY());
			        
			        
			        String Excoordinate = sheet.getRow(i+1).getCell(7).toString();
			        System.out.println("Expected x coordinate: "+Excoordinate);
			        String Eycoordinate = sheet.getRow(i+1).getCell(8).toString(); 
			        System.out.println("Expected y coordinate: "+Eycoordinate);
			        
			        if(Excoordinate.contains(".")) {
			        	Excoordinate = Excoordinate.substring(0,Excoordinate.indexOf("."));
			        }
			        if(Eycoordinate.contains(".")) {
			        	Eycoordinate = Eycoordinate.substring(0,Eycoordinate.indexOf("."));
			        }
			        
			        if (Axcoordinate.equals(Excoordinate)){
						 row.createCell(22).setCellValue("Matches");
					}else {
						 row.createCell(22).setCellValue("Not Matches");
						 highlight(cellval);
					}   
			        if (Aycoordinate.equals(Eycoordinate)){
						 row.createCell(23).setCellValue("Matches");
					 }else {
						 row.createCell(23).setCellValue("Not Matches");
						 highlight(cellval);
					 } 
			        
			        row.createCell(10).setCellValue(ATextvalue);
			        row.createCell(11).setCellValue(Afontsize);
			        row.createCell(12).setCellValue(Afontcolor);
			        row.createCell(13).setCellValue(Abackgroundcolor);
			        row.createCell(14).setCellValue(Afontfamily);
			        row.createCell(15).setCellValue(Apoint.getX());
			        row.createCell(16).setCellValue(Apoint.getY());
			} catch (NoSuchElementException e) {
			    present = false;
			    System.out.println(present);
			}
			 
			row.createCell(9).setCellValue(present);
        }
        FileOutputStream outputStream = new FileOutputStream("/Users/ramkumars/eclipse-workspace/UIComparision/src/test/java/TestData/ToolsQATestData3.xls");
        wb.write(outputStream);
        outputStream.close();
		
	}
	public void highlight(String cellval) {
		JavascriptExecutor js = (JavascriptExecutor) driver;
	    WebElement highlight= driver.findElement(By.xpath(cellval));
        js.executeScript("arguments[0].style.backgroundColor='yellow'", highlight);
	}
	
	public String appenddatetime() {
		LocalDateTime currentDateTime = LocalDateTime.now(); 
        System.out.println("Current date and time: "
                           + currentDateTime); 
        String date = currentDateTime.toString();
        date = date.replace("-", "").replace(":", "").replace(".", "");
        return date;
	}

	public void takeSnapShot(WebDriver webdriver,String fileWithPath) throws Exception{
		//Convert web driver object to TakeScreenshot
		TakesScreenshot scrShot =((TakesScreenshot)webdriver);
		//Call getScreenshotAs method to create image file
		File SrcFile=scrShot.getScreenshotAs(OutputType.FILE);
		//Move image file to new destination
		File DestFile=new File(fileWithPath);
		//Copy file at destination
		FileUtils.copyFile(SrcFile, DestFile);
		}
		



	@DataProvider(name="ToolData", parallel=false)
	public Iterator<Object[]> getExcelData() throws IOException{
		@SuppressWarnings("rawtypes")
		ArrayList<HashMap> excelData;
		TestDataReader.readExcelFile objExcelFile = new TestDataReader.readExcelFile();
		//excelData = objExcelFile.readExcel("E:\\ExcelData","ToolsQATestData.xls","Sheet1");
		excelData = objExcelFile.readExcel("/Users/ramkumars/eclipse-workspace/UIComparision/src/test/java/TestData","ToolsQATestData3.xls","Sheet1");
		
		List<Object[]> dataArray = new ArrayList<Object[]>();
		for(HashMap data : excelData){
			dataArray.add(new Object[] { data });
			}
		return dataArray.iterator();
	}

	//-------------Generate XPATH FUNCTION -----------
		public static String AbsoluteXPath(WebElement element, WebDriver driver)
		    {
		        return (String) ((JavascriptExecutor) driver).executeScript(
		                "function absoluteXPath(element) {"+
		                        "var comp, comps = [];"+
		                        "var parent = null;"+
		                        "var xpath = '';"+
		                        "var getPos = function(element) {"+
		                        "var position = 1, curNode;"+
		                        "if (element.nodeType == Node.ATTRIBUTE_NODE) {"+
		                        "return null;"+
		                        "}"+
		                        "for (curNode = element.previousSibling; curNode; curNode = curNode.previousSibling) {"+
		                        "if (curNode.nodeName == element.nodeName) {"+
		                        "++position;"+
		                        "}"+
		                        "}"+
		                        "return position;"+
		                        "};"+
		"if (element instanceof Document) {"+
		                        "return '/';"+
		                        "}"+
		"for (; element && !(element instanceof Document); element = element.nodeType == Node.ATTRIBUTE_NODE ? element.ownerElement : element.parentNode) {"+
		                        "comp = comps[comps.length] = {};"+
		                        "switch (element.nodeType) {"+
		                        "case Node.TEXT_NODE:"+
		                        "comp.name = 'text()';"+
		                        "break;"+
		                        "case Node.ATTRIBUTE_NODE:"+
		                        "comp.name = '@' + element.nodeName;"+
		                        "break;"+
		                        "case Node.PROCESSING_INSTRUCTION_NODE:"+
		                        "comp.name = 'processing-instruction()';"+
		                        "break;"+
		                        "case Node.COMMENT_NODE:"+
		                        "comp.name = 'comment()';"+
		                        "break;"+
		                        "case Node.ELEMENT_NODE:"+
		                        "comp.name = element.nodeName;"+
		                        "break;"+
		                        "}"+
		                        "comp.position = getPos(element);"+
		                        "}"+
		"for (var i = comps.length - 1; i >= 0; i--) {"+
		                        "comp = comps[i];"+
		                        "xpath += '/' + comp.name.toLowerCase();"+
		                        "if (comp.position !== null) {"+
		                        "xpath += '[' + comp.position + ']';"+
		                        "}"+
		                        "}"+
		"return xpath;"+
		"} return absoluteXPath(arguments[0]);", element);
		 }
	}


