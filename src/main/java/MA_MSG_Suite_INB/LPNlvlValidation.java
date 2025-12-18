package MA_MSG_Suite_INB;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.apache.poi.ss.usermodel.*;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.FileInputStream;
import java.time.Duration;
import java.util.*;

/**
 * =========================================================
 * LPNlvlValidation
 * ---------------------------------------------------------
 * Purpose:
 * - Validates LPNs at ASN level via WM Mobile UI
 * - Executes steps TESTCASE-wise (TST_1 → all steps → TST_2)
 * - Captures execution evidence (screenshots + doc)
 *
 * Screenshots:
 * - Captured AFTER every major step using TestcaseReporter
 * =========================================================
 */
public class LPNlvlValidation {

    private WebDriver driver;

    // ✅ Centralized Excel paths
    private static final String DATA_EXCEL_PATH = ExcelReaderIB.DATA_EXCEL_PATH;
    private static final String LOGIN_EXCEL_PATH = ExcelReaderIB.LOGIN_EXCEL_PATH;

    /**
     * Entry point called from inbound controller
     */
    public void execute(WebDriver driver) {
        try {
            // -------------------------------------------------
            // Step A: Read Excel and group rows by Testcase
            // -------------------------------------------------
            Map<String, List<Row>> testcaseMap = readTestcasesFromExcel();

            // -------------------------------------------------
            // Step B: Execute testcases sequentially
            // Order:
            //   TST_1 → all steps → TST_2 → all steps
            // -------------------------------------------------
            for (String testcaseId : testcaseMap.keySet()) {

                System.out.println("\n==============================");
                System.out.println("🚀 Executing Testcase: " + testcaseId);
                System.out.println("==============================");

                // Initialize DOCX report for this testcase
                TestcaseReporter.initTestcase(testcaseId);

                try {
                    // -------------------------------------------------
                    // Step 1: Login
                    // -------------------------------------------------
                    driver = login(testcaseId);

                    // -------------------------------------------------
                    // Step 2: Open ASN module
                    // -------------------------------------------------
                    openASNModule(testcaseId);

                    // -------------------------------------------------
                    // Step 3: Process each ASN for this testcase
                    // -------------------------------------------------
                    for (Row row : testcaseMap.get(testcaseId)) {
                        String asn = getCellValue(row.getCell(1)); // ASN column
                        if (asn.isEmpty()) continue;

                        processASN(asn, testcaseId);
                    }

                } finally {
                    // -------------------------------------------------
                    // Step 4: Cleanup browser
                    // -------------------------------------------------
                    if (driver != null) {
                        driver.quit();

                        // 📸 Screenshot note:
                        // No screenshot here because browser is already closed
                        TestcaseReporter.closeTestcase(testcaseId);
                    }
                }
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // =====================================================
    // LOGIN SECTION
    // =====================================================

    /**
     * Logs into application using Login.xlsx
     * Screenshot taken AFTER successful login
     */
    private WebDriver login(String testcaseId) throws Exception {

        WebDriverManager.chromedriver().setup();
        WebDriver driver = new ChromeDriver();
        driver.manage().window().maximize();

        Map<String, String> login = readLogin();

        // Navigate to login URL
        driver.get(login.get("LOGIN_URL"));

        // 📸 Screenshot:
        // After opening login page
        TestcaseReporter.addStep(driver, testcaseId, "Opened Login URL");

        // Enter credentials
        driver.findElement(By.id("username")).sendKeys(login.get("USERNAME"));
        driver.findElement(By.id("password")).sendKeys(login.get("PASSWORD"));
        driver.findElement(By.id("kc-login")).click();

        // Wait until landing page loads
        new WebDriverWait(driver, Duration.ofSeconds(30))
                .until(ExpectedConditions.presenceOfElementLocated(By.tagName("ion-app")));

        // 📸 Screenshot:
        // After successful login
        TestcaseReporter.addStep(driver, testcaseId, "Login successful");

        return driver;
    }

    // =====================================================
    // NAVIGATION SECTION
    // =====================================================

    /**
     * Opens ASN module from WM Mobile menu
     * Screenshot taken after menu open and ASN open
     */
    private void openASNModule(String testcaseId) throws Exception {

        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(20));

        // Open hamburger menu
        WebElement menuToggle = wait.until(
                ExpectedConditions.elementToBeClickable(
                        By.cssSelector("ion-button[data-component-id='menu-toggle-button']"))
        );
        menuToggle.click();

        // 📸 Screenshot:
        // After opening left menu
        TestcaseReporter.addStep(driver, testcaseId, "Menu opened");

        // Search for ASN
        WebElement searchBox = wait.until(
                ExpectedConditions.presenceOfElementLocated(
                        By.xpath("//input[@placeholder='Search Menu...']"))
        );
        searchBox.sendKeys("ASNS");

        // Click ASN module
        WebElement asnsBtn = wait.until(
                ExpectedConditions.elementToBeClickable(By.id("ASN"))
        );
        asnsBtn.click();

        // 📸 Screenshot:
        // After ASN module opened
        TestcaseReporter.addStep(driver, testcaseId, "ASN module opened");
    }

    // =====================================================
    // ASN → LPN VALIDATION SECTION
    // =====================================================

    /**
     * Processes a single ASN:
     * - Search ASN
     * - Open ASN card
     * - Open Related Links
     * - Navigate to LPN screen
     *
     * Screenshot taken AFTER every sub-step
     */
    private void processASN(String asn, String testcaseId) throws Exception {

        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(20));

        // Search ASN
        WebElement filterInput = wait.until(
                ExpectedConditions.elementToBeClickable(By.xpath("//input[contains(@id,'ion-input')]"))
        );
        filterInput.clear();
        filterInput.sendKeys(asn);
        filterInput.sendKeys(Keys.ENTER);

        // 📸 Screenshot:
        // After ASN search
        TestcaseReporter.addStep(driver, testcaseId, "ASN searched: " + asn);

        // Open ASN card
        WebElement asnCard = wait.until(
                ExpectedConditions.elementToBeClickable(
                        By.cssSelector("card-view[data-component-id='Card-View'] div.card-row.primary"))
        );
        asnCard.click();

        // 📸 Screenshot:
        // After ASN card opened
        TestcaseReporter.addStep(driver, testcaseId, "ASN card opened: " + asn);

        // Click Related Links
        WebElement relatedLinks = wait.until(
                ExpectedConditions.elementToBeClickable(
                        By.cssSelector("button[data-component-id='relatedLinks']"))
        );
        relatedLinks.click();

        // 📸 Screenshot:
        // After Related Links opened
        TestcaseReporter.addStep(driver, testcaseId, "Related Links opened");

        // Navigate to LPN screen
        WebElement lpnLink = wait.until(
                ExpectedConditions.elementToBeClickable(
                        By.cssSelector("ion-item[data-component-id='LPN']"))
        );
        lpnLink.click();

        // 📸 Screenshot:
        // After LPN screen opened
        TestcaseReporter.addStep(driver, testcaseId, "Navigated to LPN screen");
    }

    // =====================================================
    // EXCEL HELPERS
    // =====================================================

    /**
     * Groups rows by Testcase (TST_1, TST_2...)
     */
    private Map<String, List<Row>> readTestcasesFromExcel() throws Exception {

        Map<String, List<Row>> map = new LinkedHashMap<>();

        try (FileInputStream fis = new FileInputStream(DATA_EXCEL_PATH);
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet sheet = wb.getSheet("LPNlvlValidation");
            if (sheet == null)
                throw new RuntimeException("Sheet 'LPNlvlValidation' not found");

            Iterator<Row> it = sheet.iterator();
            it.next(); // skip header

            while (it.hasNext()) {
                Row row = it.next();
                String tc = getCellValue(row.getCell(0));

                if (!tc.matches("TST_\\d+")) continue;

                map.computeIfAbsent(tc, k -> new ArrayList<>()).add(row);
            }
        }
        return map;
    }

    /**
     * Reads login credentials from Login.xlsx
     */
    private Map<String, String> readLogin() throws Exception {

        Map<String, String> map = new HashMap<>();

        try (FileInputStream fis = new FileInputStream(LOGIN_EXCEL_PATH);
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet sheet = wb.getSheet("Login");
            Row row = sheet.getRow(1);

            map.put("LOGIN_URL", getCellValue(row.getCell(0)));
            map.put("USERNAME", getCellValue(row.getCell(2)));
            map.put("PASSWORD", getCellValue(row.getCell(3)));
        }
        return map;
    }

    /**
     * Safe Excel cell reader
     */
    private String getCellValue(Cell cell) {
        if (cell == null) return "";
        return new DataFormatter().formatCellValue(cell).trim();
    }
}
