package BankingTestCases;

import java.util.HashMap;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import BankingUtilities.ExcelRecord;
import io.github.bonigarcia.wdm.WebDriverManager;

public class Snapdeal {
	
	WebDriver driver;
	String[] aa;
	String filePathh=System.getProperty("user.dir") + "\\ExcelFolder\\ExcelRead.xlsx";
	String sheetNamee ="Sheet1";
	
	@BeforeTest
	public void dealSucc() {
		
		WebDriverManager.chromedriver().setup();
		driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.get("https://sellers.snapdeal.com/seller_diaries");
	}
	
	@Test
	public void pStories() throws Exception {
		
		Thread.sleep(4000);
		
		List<WebElement> pstories = driver.findElements(By.xpath("//*[@id='pgLogin:j_id109']/div/a"));	
		int linksCount = pstories.size();
		aa = new String[linksCount];
		
		for(int i=0;i<linksCount;i++) {			
			aa[i] = pstories.get(i).getAttribute("href");
			System.out.println("Links are.." + aa[i]);
		}
		
		for(int i=0;i<linksCount;i++) {
			driver.navigate().to(aa[i]);
			String txtVal = driver.findElement(By.xpath("//*[@itemscope='itemscope']/h2")).getText();
			System.out.println("Text are.." + txtVal); 
			ExcelRecord dde = new ExcelRecord(filePathh, sheetNamee);
			dde.writeData(txtVal);
			dde.saveAndClose();
		}
		
		
	}
}
