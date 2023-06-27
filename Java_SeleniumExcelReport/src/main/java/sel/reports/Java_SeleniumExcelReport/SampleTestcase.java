package sel.reports.Java_SeleniumExcelReport;

import java.time.Duration;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

import io.github.bonigarcia.wdm.WebDriverManager;


public class SampleTestcase {

	private static WebDriver driver;
	public static String Status;
	public static String SheetName = "TestResults";


	
	public static void main(String[] args) {
		ExcelReporter exrep = new ExcelReporter();
		exrep.startResult();
		exrep.startTestCase(SheetName);


		WebDriverManager.chromedriver().setup();
		driver = new ChromeDriver();

		driver.manage().window().maximize();
		driver.manage().deleteAllCookies();
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(3));

		//Step 1 Load Google take screenshot
		driver.get("https://www.google.com/");
		
		String title = driver.getTitle().trim();
		System.out.println(driver.getTitle());

		String tit = "Google";
		if (title.equalsIgnoreCase(tit)) {
			Status = "Pass";

		} else {
			Status = "Fail";
		}
		System.out.println(Status);
		
		long GooglePagescreenshotId = exrep.takeSnap(driver);
		exrep.reportStep("Open URL", Status, GooglePagescreenshotId);
		
		
		//Step 2 Search Amazon take screenshot
		driver.findElement(By.xpath("//textarea[@id='APjFqb']")).sendKeys("Amazon");
		driver.findElement(By.xpath("//input[@aria-label='Google Search']")).click();

		
		String SearchedPagetitle = driver.getTitle().trim();
		System.out.println(driver.getTitle());

		String searchedpagetit = "amazon - Google Search";
		if (searchedpagetit.equalsIgnoreCase(searchedpagetit)) {
			Status = "Pass";

		} else {
			Status = "Fail";
		}
		System.out.println(Status);
		
		long searchedTitscreenshotId = exrep.takeSnap(driver);
		exrep.reportStep("Page Search Verification", Status,searchedTitscreenshotId);
		
		driver.quit();
		
		exrep.endTestcase("AutomationTestReport");

	}

}
