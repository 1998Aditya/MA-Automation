package MA_MSG_Suite_INB;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.apache.poi.ss.usermodel.*;
import org.openqa.selenium.*;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.File;
import java.io.FileInputStream;
import java.time.Duration;
import java.util.*;

public class InductiLPN_MFS {

    // Updated to use centralized Excel paths
    private static final String EXCEL_PATH = ExcelReaderIB.DATA_EXCEL_PATH;
    private WebDriver driver;

    public InductiLPN_MFS(WebDriver driver) {
        this.driver = driver;
    }

    public void execute() throws InterruptedException {
        System.out.println("=== Step: Induct iLPN (MFS) Execution Started ===");

        // ✅ Step 1: Find LCIDs with ConditionCode FI or CR
        List<String> lcidList = getLCIDsForInduction();
        if (lcidList.isEmpty()) {
            System.out.println("⚠ No LCID found with ConditionCode 'FI' or 'CR'. Exiting Induct iLPN process.");
            return;
        }

        System.out.println("✅ Found LCIDs for Induction: " + lcidList);

        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(15));

        try {
            // === Step 2: Click menu toggle ===
            WebElement menuToggle = wait.until(
                    ExpectedConditions.elementToBeClickable(
                            By.cssSelector("ion-button[data-component-id='menu-toggle-button']"))
            );
            menuToggle.click();

            // === Step 3: Search for WM Mobile ===
            WebElement searchBox = wait.until(
                    ExpectedConditions.presenceOfElementLocated(
                            By.xpath("//input[@placeholder='Search Menu...']"))
            );
            searchBox.clear();
            searchBox.sendKeys("WM Mobile");

            WebElement wmMobileButton = wait.until(
                    ExpectedConditions.elementToBeClickable(
                            By.xpath("//button[@data-component-id='WMMobile']"))
            );

            // ✅ Capture the existing window handles BEFORE clicking
            Set<String> oldWindows = driver.getWindowHandles();
            wmMobileButton.click();
            System.out.println("📱 Clicked on WM Mobile button. Waiting for new tab...");

            // ✅ Wait for up to 10 seconds for a new tab to open
            String newTabHandle = waitForNewTab(driver, oldWindows, 10);

            if (newTabHandle == null) {
                System.out.println("❌ No new tab opened after clicking WM Mobile.");
                return;
            }

            driver.switchTo().window(newTabHandle);
            System.out.println("✅ Switched to WM Mobile tab.");

            // === Step 5: Transaction Name ===
            String trxnName = "JD Induct From Reject or MFS";
            System.out.println("📄 Using Transaction Name: " + trxnName);

            // === Step 6: Search for Transaction ===
            WebElement newTabSearch = wait.until(
                    ExpectedConditions.presenceOfElementLocated(
                            By.xpath("//input[@placeholder='Search']"))
            );
            newTabSearch.click();
            newTabSearch.sendKeys(trxnName);
            newTabSearch.sendKeys(Keys.ENTER);

            Thread.sleep(3000);

            // === Step 7: Click transaction by matching its text ===
            boolean clicked = false;
            List<WebElement> allElements = driver.findElements(By.xpath("//*"));
            for (WebElement el : allElements) {
                try {
                    if (el.isDisplayed() && el.getText().trim().equalsIgnoreCase(trxnName)) {
                        el.click();
                        clicked = true;
                        break;
                    }
                } catch (Exception ignored) {}
            }

            if (!clicked) {
                System.out.println("⚠ Could not find transaction: " + trxnName);
                return;
            }

            System.out.println("✅ Successfully clicked transaction: " + trxnName);

            // === Step 8: Loop through all eligible LCIDs ===
            Thread.sleep(3000);
            for (String lcid : lcidList) {
                try {
                    WebElement scanInput = wait.until(
                            ExpectedConditions.presenceOfElementLocated(
                                    By.xpath("//input[@placeholder='Scan Inbound Container']"))
                    );
                    scanInput.click();
                    scanInput.clear();
                    scanInput.sendKeys(lcid);
                    scanInput.sendKeys(Keys.ENTER);
                    System.out.println("📦 Scanned LCID: " + lcid);

                    // ✅ Wait for and click the “Ok” button in the alert popup
                    try {
                        WebElement okButton = new WebDriverWait(driver, Duration.ofSeconds(10))
                                .until(ExpectedConditions.elementToBeClickable(
                                        By.xpath("//button[.//span[text()='Ok']]")));
                        okButton.click();
                        System.out.println("✅ Clicked OK for LCID: " + lcid);
                    } catch (TimeoutException te) {
                        System.out.println("⚠ No OK button appeared for LCID: " + lcid);
                    }

                    Thread.sleep(1500); // short pause before scanning next LCID
                } catch (Exception e) {
                    System.out.println("⚠ Failed to scan LCID " + lcid + ": " + e.getMessage());
                }
            }

            System.out.println("✅ All eligible LCIDs scanned successfully.");

        } catch (Exception e) {
            e.printStackTrace();
        }

        System.out.println("=== Induct iLPN (MFS) Process Completed ===");
    }

    /**
     * ✅ Waits for a new browser tab to open within the specified timeout.
     */
    private String waitForNewTab(WebDriver driver, Set<String> oldHandles, int timeoutSeconds) {
        int waited = 0;
        while (waited < timeoutSeconds * 2) { // check every 500ms
            Set<String> newHandles = new HashSet<>(driver.getWindowHandles());
            newHandles.removeAll(oldHandles);
            if (!newHandles.isEmpty()) {
                return newHandles.iterator().next();
            }
            try {
                Thread.sleep(500);
            } catch (InterruptedException ignored) {}
            waited++;
        }
        return null;
    }

    /**
     * ✅ Returns a list of LCIDs that have Condition_Code = "FI" or "CR"
     */
    private List<String> getLCIDsForInduction() {
        List<String> lcidList = new ArrayList<>();
        // Check if caller requested filtering by Testcase (e.g. System.setProperty("testcase","TST_1"))
        boolean doFilterByTestcase = System.getProperty("testcase") != null && !System.getProperty("testcase").trim().isEmpty();
        String testcaseToRun = doFilterByTestcase ? System.getProperty("testcase").trim() : "";

        try (FileInputStream fis = new FileInputStream(new File(EXCEL_PATH));
             Workbook workbook = WorkbookFactory.create(fis)) {

            Sheet sheet = workbook.getSheet("ReportRCV");
            if (sheet == null) {
                System.out.println("❌ Sheet 'ReportRCV' not found in Excel!");
                return lcidList;
            }

            int lcidColumn = -1;
            int conditionCol = -1;
            int testcaseCol = -1;
            Row header = sheet.getRow(0);
            if (header == null) {
                System.out.println("❌ Empty 'ReportRCV' sheet (no header).");
                return lcidList;
            }

            for (int i = 0; i < header.getLastCellNum(); i++) {
                String headerName = getCellValue(header.getCell(i));
                if ("Lcid".equalsIgnoreCase(headerName)) lcidColumn = i;
                if ("Condition_Code".equalsIgnoreCase(headerName)) conditionCol = i;
                if ("Testcase".equalsIgnoreCase(headerName)) testcaseCol = i;
            }

            if (lcidColumn == -1 || conditionCol == -1) {
                System.out.println("❌ Missing 'Lcid' or 'Condition_Code' column in Excel.");
                return lcidList;
            }

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                // Testcase filtering if requested
                if (doFilterByTestcase) {
                    String tcValue = testcaseCol != -1 ? getCellValue(row.getCell(testcaseCol)) : getCellValue(row.getCell(0));
                    if (!testcaseToRun.equalsIgnoreCase(tcValue)) continue;
                }

                String lcid = getCellValue(row.getCell(lcidColumn));
                String conditionCode = getCellValue(row.getCell(conditionCol));

                if (!lcid.isEmpty() && (conditionCode.contains("FI") || conditionCode.contains("CR"))) {
                    lcidList.add(lcid);
                }
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
        return lcidList;
    }

    private String getCellValue(Cell cell) {
        if (cell == null) return "";
        DataFormatter formatter = new DataFormatter();
        return formatter.formatCellValue(cell).trim();
    }

    public static WebDriver setupDriver() {
        WebDriverManager.chromedriver().setup();
        WebDriver driver = new org.openqa.selenium.chrome.ChromeDriver();
        driver.manage().window().maximize();
        return driver;
    }
}
