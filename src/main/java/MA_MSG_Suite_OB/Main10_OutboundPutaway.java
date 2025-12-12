
package MA_MSG_Suite_OB;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStreamReader;
import java.time.Duration;








import com.google.gson.*;
import com.google.gson.stream.JsonReader;
import okhttp3.*;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

import com.google.gson.JsonArray;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;
import io.github.bonigarcia.wdm.WebDriverManager;
import okhttp3.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.FileInputStream;
import java.io.IOException;
import java.time.Duration;






public class Main10_OutboundPutaway {
    public static WebDriver driver;

    public static void main(String filePath,String testcase,String env) throws InterruptedException {
//        String filePath = "C:\\Users\\2210420\\IdeaProjects\\msg-runner\\OOdata.xlsx";
//        String testcase = "TST_001";   // input parameter
//        String env = "QA";             // environment
        closeExcelIfOpen();
        runOutboundPutawayForTestcase(filePath, "Outbound", testcase, env);
    }

    public static void runOutboundPutawayForTestcase(String filePath, String sheetName, String testcaseId, String env) throws InterruptedException {
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheet(sheetName);
            int matchCount = 0;

            for (Row row : sheet) {
                Cell testcaseCell = row.getCell(0); // first column = testcase ID
                if (testcaseCell != null) {
                    String cellValue = testcaseCell.toString().trim();
                    if (cellValue.equalsIgnoreCase(testcaseId.trim())) {

                        String status = "Failed"; // default to Failed

                        try {
                            System.out.println("\n🔁 Running " + testcaseId + " for row " + row.getRowNum());

                            // Load variables from this row
                            String pallet   = getCellValue(row.getCell(3));
                            String location = getCellValue(row.getCell(5));

                            System.out.println("Pallet: " + pallet);
                            System.out.println("Location: " + location);

                            // Setup driver
                            WebDriverManager.chromedriver().setup();
                            ChromeOptions options = new ChromeOptions();
                            options.addArguments("--start-maximized");
                            driver = new ChromeDriver(options);

                            driver.manage().window().maximize();

                            // Login and navigation
                            Main1_URL_Login1 login1 = new Main1_URL_Login1(driver, env);
                            login1.execute();
                            System.out.println("✅ Login done");

                            SearchMenuWM("WM Mobile", "WMMobile");
                            //SearchInWmMobile("JD OB Putaway To Staging", "jdobputawaytostaging");
                            Thread.sleep(5000);
                            SearchInWmMobile("JD OB Putaway To Staging","jdobputawaytostaging");
                            // Execute OutboundPutaway
                            Thread.sleep(5000);
                            OutboundPutaway(pallet, location);

                            status = "Passed"; // if no exception, mark as Passed
                        } catch (Exception e) {
                            System.err.println("❌ Error during execution: " + e.getMessage());
                            e.printStackTrace();
                        }
                        finally {
//                            if (driver != null) {
//                                driver.quit();
//                            }

                            // Write status to column 1 (second column)
                            Cell statusCell = row.getCell(1); // column index 1
                            if (statusCell == null) {
                                statusCell = row.createCell(1);
                            }
                            statusCell.setCellValue(status);
                            closeExcelIfOpen();
// ✅ Save Excel immediately after updating this row
                            try (FileOutputStream fos = new FileOutputStream(filePath)) {
                                workbook.write(fos);
                                System.out.println("📄 Excel updated for row " + row.getRowNum() + " with status: " + status);
                            } catch (Exception e) {
                                System.err.println("❌ Error saving Excel after row update: " + e.getMessage());
                            }

                        }





                    }
                }
            }

            if (matchCount == 0) {
                System.out.println("❌ No matching rows found for testcase: " + testcaseId);
            }
            closeExcelIfOpen();
            try (FileOutputStream fos = new FileOutputStream(filePath)) {
                workbook.write(fos);
                System.out.println("✅ Excel updated with status.");
            } catch (Exception e) {
                System.err.println("❌ Error saving Excel: " + e.getMessage());
            }


        } catch (Exception e) {
            System.err.println("❌ Error processing Excel: " + e.getMessage());
            e.printStackTrace();
        }
    }

    public static void OutboundPutaway(String pallet, String location) throws InterruptedException {
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
        JavascriptExecutor js = (JavascriptExecutor) driver;

        try {

            WebElement palletInput = wait.until(ExpectedConditions.elementToBeClickable(
                    By.cssSelector("input[data-component-id='acceptcontainer_barcodetextfield_scancontainer']")));


            palletInput.clear();
            palletInput.sendKeys(pallet);
            palletInput.sendKeys(Keys.ENTER);
            System.out.println("Entered pallet: " + pallet);

            Thread.sleep(2000);

            WebElement locationInput = wait.until(ExpectedConditions.elementToBeClickable(
                    By.cssSelector("input[data-component-id='acceptlocation_barcodetextfield_scanlocation']")));
            locationInput.clear();
            locationInput.sendKeys(location);
            locationInput.sendKeys(Keys.ENTER);
            System.out.println("Entered location: " + location);

        } catch (Exception e) {
            System.err.println("❌ Error in OutboundPutaway: " + e.getMessage());
        }
    }

    private static String getCellValue(Cell cell) {
        return cell == null ? "" : cell.toString().trim();
    }
    public static void SearchMenuWM(String Keyword, String id)  {
        WebDriverWait wait1 = new WebDriverWait(driver, Duration.ofSeconds(15));
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
        JavascriptExecutor js = (JavascriptExecutor) driver;




        int maxRetries = 6; // Try up to 2 times (1 initial + 1 retry)
        for (int attempt = 1; attempt <= maxRetries; attempt++) {
            System.out.println("⏳ Waiting 10 seconds before refreshing...");

            try {
                Thread.sleep(10000); // Wait 1 minute
            } catch (InterruptedException e) {
                throw new RuntimeException(e);
            }

//                // 🔄 Click the refresh button inside shadow DOM
//                WebElement refreshHost = driver.findElement(By.cssSelector("ion-button.refresh-button"));
//                js = (JavascriptExecutor) driver;
//                WebElement refreshButton = (WebElement) js.executeScript(
//                        "return arguments[0].shadowRoot.querySelector('button.button-native')", refreshHost);
//                refreshButton.click();
//                System.out.println("🔄 Refresh button clicked.");

            // Locate using data-component-id
            WebElement refreshBtn = wait.until(
                    ExpectedConditions.elementToBeClickable(
                            By.cssSelector("ion-button[data-component-id='refresh']")
                    )
            );

            // Click the button
            refreshBtn.click();

            // Optional: verify action or add logging
            System.out.println("Refresh button clicked successfully!");

            // Optional: wait for UI to settle
            try {
                Thread.sleep(3000);
            } catch (InterruptedException e) {
                throw new RuntimeException(e);
            }

            try {
                WebElement shadowHost = wait1.until(ExpectedConditions.presenceOfElementLocated(
                        By.cssSelector("ion-button[data-component-id='menu-toggle-button']")
                ));
                SearchContext shadowRoot = (SearchContext) js.executeScript("return arguments[0].shadowRoot", shadowHost);
                WebElement nativeButton = shadowRoot.findElement(By.cssSelector("button.button-native"));

// wait for overlay to disappear
                wait1.until(ExpectedConditions.invisibilityOfElementLocated(By.cssSelector("manh-overlay-container")));

// click via JS to avoid interception
                js.executeScript("arguments[0].click();", nativeButton);

                System.out.println("Menu toggle button clicked.");




                break;
            }catch (Exception e) {
                System.err.println("Error: " + e.getMessage());
                e.printStackTrace(System.err);


            }


        }




        try {
            // Locate the inner input directly under ion-input
            WebElement innerInput = wait.until(ExpectedConditions.presenceOfElementLocated(
                    By.cssSelector("ion-input[data-component-id='search-input'] input.native-input")
            ));

            wait.until(ExpectedConditions.elementToBeClickable(innerInput));

            innerInput.clear();
            innerInput.sendKeys(Keyword);
            System.out.println("✅ Search Done: " + Keyword);

        } catch (Exception e) {
            System.err.println("❌ Error interacting with search input: " + e.getMessage());
            e.printStackTrace();
        }
        try {


            // Wait for the button to be present and visible
            WebElement element = wait.until(ExpectedConditions.presenceOfElementLocated(
                    By.cssSelector("button#wmMobile[data-component-id=" + id + "]")
            ));


            ((JavascriptExecutor) driver).executeScript(
                    "arguments[0].scrollIntoView({block: 'center'});", element
            );

            ((JavascriptExecutor) driver).executeScript(
                    "arguments[0].click();", element
            );

            System.out.println("Clicked element with id: " + id);
            ArrayList<String> tabs = new ArrayList<>(driver.getWindowHandles());
            driver.switchTo().window(tabs.get(1));

        } catch (Exception e) {

            System.err.println("Failed to click element with id: " + id);
            e.printStackTrace();

        }


    }
    static void SearchInWmMobile(String Transaction, String ComponentId)  {
        JavascriptExecutor js = (JavascriptExecutor) driver;
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));


        try {


            WebElement searchInput = wait.until(ExpectedConditions.elementToBeClickable(
                    By.cssSelector("ion-searchbar[data-component-id='search'] input[type='search']")));

            // Clear any existing text
            searchInput.clear();

            // Type the search text
            searchInput.sendKeys(Transaction);

            // Optionally, press ENTER if the search requires submission
            searchInput.sendKeys(Keys.ENTER);


//            WebElement searchInput1 = wait.until(ExpectedConditions.elementToBeClickable(
//                    By.cssSelector("input.searchbar-input[placeholder='Search']")));
////                    By.xpath("//input[@type='search' and @placeholder='Search']")));
//            searchInput1.click();
//            searchInput1.clear();
//            Thread.sleep(3000);
//            searchInput1.sendKeys(Transaction);
        } catch (Exception e) {
            System.err.println("❌ Error in "+Transaction + e.getMessage());
            e.printStackTrace();
        }
        // Locate the ion-label using its data-component-id

                WebElement labelElement = wait.until(ExpectedConditions.elementToBeClickable(
                By.cssSelector("ion-label[data-component-id='" + ComponentId + "']")
        ));

        // Scroll into view to ensure it's interactable
        ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", labelElement);

        // Click using JavaScript (in case native click doesn't work)
        ((JavascriptExecutor) driver).executeScript("arguments[0].click();", labelElement);

        System.out.println("Clicked on" +Transaction+" label.");


    }
    private static void closeExcelIfOpen() {
        try {
            Process process = Runtime.getRuntime().exec("tasklist");
            BufferedReader reader = new BufferedReader(new InputStreamReader(process.getInputStream()));
            String line;
            boolean excelRunning = false;
            while ((line = reader.readLine()) != null) {
                if (line.toLowerCase().contains("excel.exe")) {
                    excelRunning = true;
                    break;
                }
            }
            if (excelRunning) {
                System.out.println("⚠️ Excel is open. Closing it...");
                Runtime.getRuntime().exec("taskkill /IM excel.exe /F");
                Thread.sleep(2000);
            }
        } catch (Exception e) {
            System.err.println("⚠️ Could not check/close Excel: " + e.getMessage());
        }
    }






}

