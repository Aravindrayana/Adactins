package com.baseclass;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;

public class BaseClass {
	public static WebDriver driver;

	public static WebDriver getBrowser(String web) {

		if (web.equalsIgnoreCase("chrome")) {
			System.setProperty("webdriver.chrome.driver",
					System.getProperty("user.dir") + "\\Driver\\chromedriver.exe");

			driver = new ChromeDriver();
		} else if (web.equalsIgnoreCase("firefox")) {
			System.setProperty("webdriver.gecko.driver", System.getProperty("user.dir") + "\\Driver\\chromedriver.exe");
			driver = new FirefoxDriver();
		}
		driver.manage().window().maximize();
		return driver;
	}

	public static void click(WebElement element) {
		element.click();
	}

	public static void sendKeys(WebElement element, String value) {

		element.sendKeys(value);

	}

	public static void getText(WebElement element) {

		String text = element.getText();
		System.out.println(text);
	}

	public static void getUrl(String Url) {

		driver.get(Url);
	}

	public static void getCurrentUrl() {

		String currentUrl = driver.getCurrentUrl();
		System.out.println(currentUrl);
	}

	public static void getTitle() {

		String title = driver.getTitle();
		System.out.println(title);
	}

	public static void displayed(WebElement element) {

		boolean displayed = element.isDisplayed();
		System.out.println(displayed);
	}

	public static void enabled(WebElement element) {

		boolean enabled = element.isEnabled();
		System.out.println(enabled);
	}

	public static void selected(WebElement element) {

		boolean selected = element.isSelected();
		System.out.println(selected);
	}

	public static void navigateTo(String Url) {

		driver.navigate().to(Url);
	}

	public static void back() {

		driver.navigate().back();
	}

	public static void forward() {

		driver.navigate().forward();
	}

	public static void refresh() {

		driver.navigate().refresh();
	}

	public static void close() {

		driver.close();
	}

	public static void quit() {

		driver.quit();
	}

	public static void Screenshot(String folder, String name) throws Throwable {

		TakesScreenshot scrnsot = (TakesScreenshot) driver;
		File fl = scrnsot.getScreenshotAs(OutputType.FILE);
		File fil = new File(System.getProperty("user.dir") + folder + name + ".png");
		FileUtils.copyFile(fl, fil);
	}

	public static void implicitwait(int num) {

		driver.manage().timeouts().implicitlyWait(num, TimeUnit.SECONDS);
	}

	public static void singleDropdown(WebElement element, String att, String value) {

		try {
			Select sc = new Select(element);

			if (att.equalsIgnoreCase("value")) {

				sc.selectByValue(value);

			} else if (att.equalsIgnoreCase("text")) {
				sc.selectByVisibleText(value);

			} else if (att.equalsIgnoreCase("index")) {

				int parseInt = Integer.parseInt(value);
				sc.selectByIndex(parseInt);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	public static void mouseover(WebElement element, String type) {

		Actions a = new Actions(driver);

		if (type.equalsIgnoreCase("move")) {

			a.moveToElement(element).build().perform();

		}

		else if (type.equalsIgnoreCase("click")) {

			a.click(element).build().perform();

		}

		else if (type.equalsIgnoreCase("rightclick")) {

			a.contextClick(element).build().perform();

		}

	}

	public static void robot(WebElement element, String rob) throws AWTException {

		Robot rc = new Robot();

		if (rob.equalsIgnoreCase("down")) {

			rc.keyPress(KeyEvent.VK_DOWN);
			rc.keyRelease(KeyEvent.VK_DOWN);

		} else if (rob.equalsIgnoreCase("enter")) {

			rc.keyPress(KeyEvent.VK_ENTER);
			rc.keyRelease(KeyEvent.VK_ENTER);

		}

	}

	public static void frames(WebElement element) {

		driver.switchTo().frame(element);

	}

	public static void Scrolldown(WebElement element) {

		JavascriptExecutor js = (JavascriptExecutor) driver;

		js.executeScript("arguments[0].scrollIntoView();", element);

	}

	public static void particularData(String path, int sheetindex, int rowindex, int cellindex) throws Throwable {
		
		File f = new File(path);
		FileInputStream fil = new FileInputStream(f);
		Workbook wb = new XSSFWorkbook(fil);
		Sheet sAt = wb.getSheetAt(sheetindex);
		Row r = sAt.getRow(rowindex);
		Cell c = r.getCell(cellindex);
		CellType cellType = c.getCellType();
		
		if (cellType.equals(CellType.STRING)) {
			
			String stringCV = c.getStringCellValue();
			System.out.println(stringCV);
			
			
		}
		else if (cellType.equals(CellType.NUMERIC)) {
			
			double numericCV = c.getNumericCellValue();
			int num = (int) numericCV;
			System.out.println(num);
		}
			
		}

	public static void All_Data(String path, int sheetindex) throws Throwable {

		File f = new File(path);
		FileInputStream fil = new FileInputStream(f);
		Workbook wb = new XSSFWorkbook(fil);
		Sheet sAt = wb.getSheetAt(sheetindex);
		// getPhysicalNumberOfRows
		int Row_Size = sAt.getPhysicalNumberOfRows();
		for (int i = 0; i < Row_Size; i++) {
			Row row = sAt.getRow(i);
			// getPhysicalNumberOfCells
			int Cell_Size = row.getPhysicalNumberOfCells();
			for (int j = 0; j < Cell_Size; j++) {
				Cell cll = row.getCell(j);
				// getCellType
				CellType cellType = cll.getCellType();
				if (cellType.equals(CellType.STRING)) {
					// getStringCellValue
					String stringCellValue = cll.getStringCellValue();
					System.out.println(stringCellValue);
				} else if (cellType.equals(CellType.NUMERIC)) {
					// getNumericCellValue
					double numericCellValue = cll.getNumericCellValue();
					int value = (int) numericCellValue;
					System.out.println(value);
				}
			}
		}

	}

}