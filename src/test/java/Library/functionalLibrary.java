package Library;

import java.util.HashMap;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.Select;

import Repository.pageElements;

public class functionalLibrary extends pageElements{
	WebDriver driver;
	HashMap<String, String> data;
	
	public functionalLibrary(WebDriver driver,	HashMap<String, String> data) {
		// TODO Auto-generated constructor stub
		super();
		this.driver = driver;
		this.data = data;
	}
	public void customclick(WebElement passelement) throws InterruptedException{
		Thread.sleep(1500);
		passelement.click();
		//submit.click();
	}
	public void customselect(WebElement passelement, String datavalue) {
		Select se = new Select(passelement);
		se.selectByVisibleText(datavalue);
	}
	public void customsettext(WebElement passelement, String datavalue) {
		passelement.sendKeys(datavalue);
	}
}
