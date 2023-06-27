package sel.reports.Java_SeleniumExcelReport;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.IndexedColors;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebDriverException;
import org.openqa.selenium.io.FileHandler;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.common.usermodel.HyperlinkType;

public class ExcelReporter {
	private static String currentDirectory, reportFileSavePath;
	private static LocalDateTime currentDateTime;
	private static String TDformat;
	public static File pathCheck;

	private static XSSFWorkbook workbook;
	private static XSSFSheet sheet;
	private static CreationHelper createHelper;

	// Get Date and Time
	public static String GetCurrentDateAndTime()
	{
		try 
		{
			currentDateTime = LocalDateTime.now();
			TDformat = currentDateTime.format(DateTimeFormatter.ofPattern("d-M-yyyy_hh:mm:ss_a")).toString();
		} 
		catch (Exception ex) 
		{
			System.out.println("Error in fetching Current Date and Time : " + ex.getMessage());
		}
		return TDformat;
	}

	// ExcelReport File storing Path
	public static String ExcelReportFileSavingPath() {
		currentDirectory = System.getProperty("user.dir");
		reportFileSavePath = currentDirectory + "\\ExcelReports\\";
		System.out.println(reportFileSavePath);
		PathCheck();
		return reportFileSavePath;
	}

	// Path Check Exists Verification
	public static void PathCheck() {
		pathCheck = new File(reportFileSavePath.toString());
		try 
		{
			if (pathCheck.exists())
			{
				File filesList[] = pathCheck.listFiles();
				for (File file : filesList) 
				{
					if (file.isFile())
					{
						file.delete();
					}
					System.out.println("Deleted");
				}
				pathCheck.mkdir();
			} else {
				pathCheck.mkdir();
			}
		} 
		catch (Exception ex) 
		{
			System.out.println("Deleting Files Error :" + ex.getMessage());

		}
	}

	// Use @BeforeSuite ==> startResult();
	// Report Config Setup
	public void startResult() {

	}

	public ExcelReporter startTestCase(String testCaseName) {

		workbook = new XSSFWorkbook();
		createHelper = workbook.getCreationHelper();
		sheet = workbook.createSheet(testCaseName);
		XSSFRow header = sheet.createRow(0);

		XSSFCellStyle style = workbook.createCellStyle();
		style.setFillBackgroundColor(IndexedColors.AQUA.getIndex());
//		style.setFillPattern(CellStyle);

		XSSFCell cell0 = header.createCell(0);
		cell0.setCellValue("S.No");
		cell0.setCellStyle(style);

		XSSFCell cell1 = header.createCell(1);
		cell1.setCellValue("Step Details");
		cell1.setCellStyle(style);

		XSSFCell cell2 = header.createCell(2);
		cell2.setCellValue("Status");
		cell2.setCellStyle(style);

		XSSFCell cell3 = header.createCell(3);
		cell3.setCellValue("Snapshot");
		cell3.setCellStyle(style);

		return this;

	}

	public long takeSnap(WebDriver driver) {
		ExcelReportFileSavingPath();
		long number = (long) Math.floor(Math.random() * 900000000L) + 10000000L;
		try 
		{
			FileHandler.copy(((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE),
					new File(reportFileSavePath + "\\images\\" + number  + ".jpg"));
		} 
		catch (Exception e) 
		{
			System.out.println("Error in capturing Screenshot");
		}

		return number;
	}

	public void reportStep(String desc, String status, long snapNumber) {
		try
		{
		XSSFRow nextRow = sheet.createRow(sheet.getLastRowNum() + 1);
		nextRow.createCell(0).setCellValue(sheet.getLastRowNum());
		nextRow.createCell(1).setCellValue(desc);
		nextRow.createCell(2).setCellValue(status);
		XSSFCell cell = nextRow.createCell(3);
		cell.setCellValue("Click here");
		//Attach Screenshots as link
		XSSFHyperlink link = (XSSFHyperlink) createHelper.createHyperlink(HyperlinkType.URL);
//				.createHyperlink();
		link.setAddress(reportFileSavePath + "\\images\\" + snapNumber + ".jpg");
		cell.setHyperlink(link);
		}
		catch(Exception ex)
		{
		System.out.println(ex.getMessage());
		}
	}

	public void endResult() {

	}

	public void endTestcase(String fileName) {
		GetCurrentDateAndTime();
		ExcelReportFileSavingPath();
		try {
			for (int i = 0; i < 3; i++)
				sheet.autoSizeColumn(i);
			workbook.write(new FileOutputStream(new File(reportFileSavePath+fileName + " - " + new Date().getTime() + ".xlsx")));
		} catch (Exception e) {
			System.out.println(e.getMessage());
			e.printStackTrace();
		}
	}

}
