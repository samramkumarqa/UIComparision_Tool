package Pages;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.stream.Collectors;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.Select;

import Library.functionalLibrary;
import Repository.pageElements;

public class searchpage extends functionalLibrary{
	WebDriver driver;
	HashMap<String, String> data;
	ArrayList<String> primarylinks;
	ArrayList<String> secondrylinks;
	ArrayList<String> thirdlinks;
	ArrayList<String> fourlinks;
	ArrayList<String> otherlinks;
	/*
	@FindBy(xpath = "//input[@id='nav-search-submit-button']")
	WebElement submit;
	@FindBy(xpath = "//select[@id='searchDropdownBox']")
	WebElement searchdrpdwn;
	@FindBy(xpath = "//input[@id='twotabsearchtextbox']")
	WebElement searchbx;
	*/
	
	
	public searchpage(WebDriver driver, HashMap<String, String> data, ArrayList<String> primarylinks, ArrayList<String> secondrylinks, ArrayList<String> thirdlinks, ArrayList<String> fourlinks, ArrayList<String> otherlinks) {
		// TODO Auto-generated constructor stub
		super(driver, data);
		this.driver = driver;
		this.data = data;
		this.primarylinks = primarylinks;
		this.secondrylinks = secondrylinks;
		this.thirdlinks = thirdlinks;
		this.fourlinks = fourlinks;
		this.otherlinks = otherlinks;
		PageFactory.initElements(driver, this);
	}
	
	public void searchText() throws IOException, InterruptedException
	{
		//Select se = new Select(driver.findElement(By.xpath("//select[@id='searchDropdownBox']")));
		//Select se = new Select(searchdrpdwn);
		//se.selectByVisibleText(data.get("Value2"));
		//searchbx.sendKeys(data.get("Value1"));
		//driver.findElement(By.xpath("//input[@id='twotabsearchtextbox']")).sendKeys(data.get("Value1"));
		//driver.findElement(By.xpath("//input[@id='nav-search-submit-button']")).click();
		//submit.click();
		customselect(searchdrpdwn, data.get("Value2"));
		customsettext(searchbx, data.get("Value1"));
		customclick(submit);
	}
	
	public void writeResults(List<Object> collectlinks, String path, String sheetname1) throws IOException {
		File file =    new File(path);
	      FileInputStream inputStream = new FileInputStream(file);
	      HSSFWorkbook wb=new HSSFWorkbook(inputStream);
	      HSSFSheet sheet=wb.getSheet(sheetname1);

	      for (int i = 1; i<=collectlinks.size()-1; i = i+1) {
	    	  String writelinks = collectlinks.get(i).toString();
	    	  System.out.println("Write in Excel : "+writelinks);
	    	  HSSFCell cell = sheet.createRow(i).createCell(2);
	    	  cell.setCellValue(writelinks);
	    	  //cell.setCellValue("writeprimarylinks"+i);
	      }
	      inputStream.close();
	      
	      FileOutputStream fos = new FileOutputStream(file);
	      wb.write(fos);
	      fos.close();
	}
	
	public void spider(String passlinks) throws InterruptedException {
		//driver.get("https://www.amazon.in/deals");
		driver.get(passlinks);
		System.out.println("Current page : "+passlinks);
		Thread.sleep(500);
		java.util.List<WebElement> links = driver.findElements(By.tagName("a"));
	      System.out.println("Number of links in the current page is " + links.size());
	     
	      for (int i = 1; i<=links.size()-1; i = i+1) {
	         
	         String collectlinks = links.get(i).getAttribute("href");
	         //System.out.println("Name of Link# " + i + links.get(i).getAttribute("href"));
	         if (collectlinks != null && !collectlinks.isEmpty() && collectlinks.contains(data.get("Domain"))) {
	        	 String getlinks="";
	        	 if(collectlinks.contains("?")) {
	        		 //System.out.println(i+" "+collectlinks.substring(0, collectlinks.indexOf("?"))); 
	        		 getlinks = collectlinks.substring(0, collectlinks.indexOf("?"));
	        	 }else if(collectlinks.contains("/ref=")){
	        		//System.out.println(i+" "+collectlinks.substring(0, collectlinks.indexOf("ref"))); 
	        		 getlinks = collectlinks.substring(0, collectlinks.indexOf("/ref="));
	        	 }else {
	        		//System.out.println(i+" "+collectlinks); 
	        		 getlinks = collectlinks;
	        	 }
	        	 
	        	 //Separate links as primary secondry and third
	        	 int count = StringUtils.countMatches(getlinks, "/");
	        	 //System.out.println("Counts : "+count);
	        	 //System.out.println("Get Links"+getlinks);
	        	 if(count==3) {
	        		 primarylinks.add(getlinks);	        		  
                 }else if(count==4) {
	        		 secondrylinks.add(getlinks);
	        	 }else if(count==5) {
	        		 thirdlinks.add(getlinks);
	        	 }else if(count==6) {
	        		 fourlinks.add(getlinks);
	        	 }else if(count>6) {
	        		 otherlinks.add(getlinks);
	        	 }
	        	
	         }
	       
	      }
	}
}
