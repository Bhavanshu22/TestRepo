package Nilesh;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.NoSuchElementException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.StaleElementReferenceException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class Random {
	public static void main(String[] args) {
		// Set ChromeDriver path (Update this path)
		System.setProperty("webdriver.chrome.driver", "D:\\My Project\\Workspace\\Data_Extract\\chromedriver.exe");

		// Configure Chrome options
		ChromeOptions options = new ChromeOptions();
//        options.addArguments("--headless"); // Run without UI
//        options.addArguments("--disable-gpu");
//        options.addArguments("--no-sandbox");

		// Initialize WebDriver
		WebDriver driver = new ChromeDriver(options);
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		driver.manage().window().maximize();

		// Base URL
		String baseUrl = "https://app.virtubox.io/bharat-tex/exhibitor-directory-website";
		driver.get(baseUrl);
		// Wait for page to load
		try {
			Thread.sleep(5000);
		} catch (InterruptedException e) {
			e.printStackTrace();
		}

		JavascriptExecutor jse = (JavascriptExecutor) driver;
		jse.executeScript("window.scrollBy(0,500)");

		// List to store data
		List<List<String>> data = new ArrayList<>();

		// Scrape all pages
		scrapeAllPages(driver, data);
		// Close the browser
//        driver.quit();
		// Save to Excel
		saveToExcel(data);
		System.out.println("Scraping completed! Data saved to exhibitor_data.xlsx");
	}

	// Function to scrape data from a single page
	private static void scrapePage(WebDriver driver, List<List<String>> data) {
		try {
			WebElement table = driver.findElement(By.cssSelector("table.vb-data-table-page3"));
			List<WebElement> rows = table.findElements(By.cssSelector("tbody tr"));

			for (WebElement row : rows) {
				List<WebElement> columns = row.findElements(By.tagName("td"));
				if (columns.size() >= 10) {
					List<String> rowData = new ArrayList<>();
					for (int i = 0; i < 10; i++) {
						rowData.add(columns.get(i).getText().trim());
					}
					data.add(rowData);
				}
			}
		} catch (NoSuchElementException e) {
			System.out.println("Table not found or no more data available.");
		}
	}

	// Function to handle pagination
	private static void scrapeAllPages(WebDriver driver, List<List<String>> data) {
        while (true) {
            scrapePage(driver, data);
            try {
                WebElement nextButton = driver.findElement(By.xpath("//a[text()='Next']"));
                if (nextButton.getAttribute("class").contains("disabled")) {
                    break; // Stop if no next page
                }
                
//                Thread.sleep(4000);
                driver.findElement(By.xpath("//button[@id='js-cookie-consent-agree']")).click();
                
//                List<WebElement> paginationList = driver.findElements(By.xpath("//ul[@class='pagination']/li"));
//                
//                for(int i=0;i<paginationList.size();i++)
//                {
//                	
//                	paginationList.get(i).click();
////                	nextButton.click();
//                	Thread.sleep(5000); // Wait for page load
////                	break;
//                }
                
                WebDriverWait wait = new WebDriverWait(driver, 10);
             // Loop through pagination elements with retry mechanism
                for (int i = 0; i < 3; i++) {  // Retry up to 3 times if stale element error occurs
                    try {
                        List<WebElement> paginationList = driver.findElements(By.xpath("//ul[@class='pagination']/li"));
                        
                        for (int j = 0; j < paginationList.size(); j++) {
                            // Re-fetch pagination elements inside the loop to avoid stale references
                            paginationList = driver.findElements(By.xpath("//ul[@class='pagination']/li"));
                            
                            WebElement page = paginationList.get(j);
                            wait.until(ExpectedConditions.elementToBeClickable(page));  // Ensure element is clickable
                            page.click();
                            
                            Thread.sleep(5000); // Wait for page load
                        }
                        break; // Exit loop if successful
                    } catch (StaleElementReferenceException e) {
                        System.out.println("Stale element reference, retrying... Attempt " + (i + 1));
                    }
                }
                        
                
            } catch (Exception e) {
                System.out.println("Pagination error: " + e.getMessage());
                break;
            }
        }
    }

	// Function to save data to Excel
	private static void saveToExcel(List<List<String>> data) {
		Workbook workbook = new XSSFWorkbook();
		Sheet sheet = workbook.createSheet("Exhibitors");
		String[] headers = { "Company Name", "Hall No.", "Booth No.", "Product Group", "Product Zone", "Contact Person",
				"Email", "Mobile", "City", "State" };
		Row headerRow = sheet.createRow(0);
		for (int i = 0; i < headers.length; i++) {
			headerRow.createCell(i).setCellValue(headers[i]);
		}

		int rowNum = 1;
		for (List<String> rowData : data) {
			Row row = sheet.createRow(rowNum++);
			for (int i = 0; i < rowData.size(); i++) {
				row.createCell(i).setCellValue(rowData.get(i));
			}
		}

		try (FileOutputStream fileOut = new FileOutputStream("exhibitor_data.xlsx")) {
			workbook.write(fileOut);
			workbook.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

}
