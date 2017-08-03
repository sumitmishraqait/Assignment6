package assignmnet6.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;


public class TestExel {
	
	//Locators
	private By byEmail = By.id("identifierId");
	private By byNext  =   By.xpath(".//*[@id='identifierNext']/content/span");
	private By byPassword = By.xpath(".//*[@id='password']/div[1]/div/div[1]/input");
	private By bySubmit = By.xpath(".//*[@id='passwordNext']/content/span");
    private By byCompose = By.xpath(".//*[text()='COMPOSE']");
    private By byTo =By.xpath(".//textarea[@name='to']");
    private By bySubject =By.xpath(".//input[@name='subjectbox']");
    private By byBody =By.xpath(".//div[@aria-label='Message Body']");
    private By bySend =By.xpath(".//*[text()='Send']");
    private By bylogout=By.cssSelector(".gb_8a.gbii");
    private By bysignout=By.xpath(".//*[text()='Sign out']");
	

	@Test(dataProvider="empLogin")
	public void VerifyInvalidLogin(String userName, String password,String To,String subject,String body) throws InterruptedException {
		System.setProperty("webdriver.chrome.driver", "C:\\Users\\sumitmishra.QAIT\\Desktop\\Test\\src\\test\\resource\\chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		//driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);
		// driver.manage().deleteAllCookies();
		  driver.get("http://gmail.com/");
		driver.findElement(byEmail).sendKeys(userName);
		driver.findElement(byNext).click();
		 Thread.sleep(2000);
		driver.findElement(byPassword).sendKeys(password);
		driver.findElement(bySubmit).click();
		 driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);
		driver.findElement(byCompose).click();
		 driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
		driver.findElement(byTo).sendKeys(To);
		driver.findElement(bySubject).sendKeys(subject);
		driver.findElement(byBody).sendKeys(body);
		driver.findElement(bySend).click();
		driver.findElement(bylogout).click();
		driver.findElement(bysignout).click();
	
		
	}
	
	@DataProvider(name="empLogin",parallel=true)
	public Object[][] loginData() {
		Object[][] arrayObject = getExcelData("C:\Users\sumitmishra.QAIT\Desktop\Test\src\test\resource","Sheet1");
		return arrayObject;
	}

	/**
	 * @param File Name
	 * @param Sheet Name
	 * @return
	 */
	public String[][] getExcelData(String fileName, String sheetName) {
		String[][] arrayExcelData = null;
		try {
			FileInputStream fs = new FileInputStream(fileName);
			Workbook wb = Workbook.getWorkbook(fs);
			Sheet sh = wb.getSheet(sheetName);

			int totalNoOfCols = sh.getColumns();
			int totalNoOfRows = sh.getRows();
			
			arrayExcelData = new String[totalNoOfRows-1][totalNoOfCols];
			
			for (int i= 1 ; i < totalNoOfRows; i++) {

				for (int j=0; j < totalNoOfCols; j++) {
					arrayExcelData[i-1][j] = sh.getCell(j, i).getContents();
				}

			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
			e.printStackTrace();
		} catch (BiffException e) {
			e.printStackTrace();
		}
		return arrayExcelData;
	}

	
}