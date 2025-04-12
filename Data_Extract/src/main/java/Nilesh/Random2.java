package Nilesh;

import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class Random2 {

	public static void main(String[] args) throws IOException
	{
	        System.setProperty("webdriver.chrome.driver", "D:\\My Project\\Workspace\\Data_Extract\\chromedriver.exe");
	        ChromeOptions options = new ChromeOptions();
	        options.addArguments("--headless");
	        options.addArguments("--disable-gpu");
	        options.addArguments("--no-sandbox");
	        WebDriver driver = new ChromeDriver(options);
	        WebDriverWait wait = new WebDriverWait(driver, 30);
	        String baseUrl = "https://app.virtubox.io/bharat-tex/exhibitor-directory-website";
	        driver.get(baseUrl);
	        List<String[]> data = new ArrayList<>();
	        data.add(new String[]{"Company Name", "Hall No.", "Booth No.", "Product Group", "Product Zone", "Contact Person", "Email", "Mobile", "City", "State"});
	        while (true) {
	            try {
	                List<WebElement> rows = driver.findElements(By.cssSelector("table.vb-data-table-page3 tbody tr"));
	                for (WebElement row : rows) {
	                    List<WebElement> columns = row.findElements(By.tagName("td"));
	                    if (columns.size() >= 10) {
	                        String[] rowData = new String[10];
	                        for (int i = 0; i < 10; i++) {
	                            rowData[i] = columns.get(i).getText().trim();
	                        }
	                        data.add(rowData);
	                    }
	                }
	                WebElement nextButton = driver.findElement(By.cssSelector("a[aria-label='Next']"));
	                if (!nextButton.isDisplayed() || nextButton.getAttribute("class").contains("disabled")) {
	                    break;
	                }
	                nextButton.click();
	                wait.until(ExpectedConditions.stalenessOf(rows.get(0))); 
	            } catch (Exception e) {
	                System.out.println("Error navigating pages: " + e.getMessage());
	                break;
	            }
	        }
	        driver.quit();
	        try (Workbook workbook = new XSSFWorkbook(); FileOutputStream fileOut = new FileOutputStream("exhibitor_data.xlsx")) {
	            Sheet sheet = workbook.createSheet("Exhibitors");
	            int rowNum = 0;
	            for (String[] rowData : data) {
	                Row row = sheet.createRow(rowNum++);
	                for (int i = 0; i < rowData.length; i++) {
	                    row.createCell(i).setCellValue(rowData[i]);
	                }
	            }
	            workbook.write(fileOut);
	        }
	        System.out.println("Scraping completed! Data saved to 'exhibitor_data.xlsx'.");
	    }
		


}
