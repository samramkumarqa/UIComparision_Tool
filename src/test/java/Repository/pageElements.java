package Repository;

import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;

public class pageElements {

	@FindBy(xpath = "//input[@id='nav-search-submit-button']")
	protected WebElement submit;
	@FindBy(xpath = "//select[@id='searchDropdownBox']")
	protected WebElement searchdrpdwn;
	@FindBy(xpath = "//input[@id='twotabsearchtextbox']")
	protected WebElement searchbx;
}
