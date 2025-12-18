package MA_MSG_Suite_INB;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.apache.poi.ss.usermodel.*;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import MA_MSG_Suite_INB.TestcaseReporter;


import java.io.FileInputStream;
import java.time.Duration;
import java.util.*;

/**
 * =========================================================
 * ItemlvlValidation
 * ---------------------------------------------------------
 * Purpose:
 * - Validate Item-level receiving for ASN
 * - Executes testcase-wise:
 *      TST_1 -> all steps
 *      TST_2 -> all steps
 * - Captures screenshots after each major step
 *
 * Evidence:
 * - Uses TestcaseReporter to generate one DOCX per testcase
 * =========================================================
 */
public class ItemlvlValidation {

    private WebDriver driver;

    // ✅ Centralized Excel paths
    private static final String DATA_EXCEL_PATH = ExcelReaderIB.DATA_EXCEL_PATH;
    private static final String LOGIN_EXCEL_PATH = ExcelReaderIB.LOGIN_EXCEL_PATH;

    /**
     * Entry point (called from controller)
     */
    public void execute(WebDriver driver) {
        try {
            // -------------------------------------------------
            // STEP A: Group Excel rows by Testcase (TST_*)
            // -------------------------------------------------
            Map<String, List<Row>> testcaseMap = readTestcasesFromExcel();

            // -------------------------------------------------
            // STEP B: Execute each testcase sequentially
            // -------------------------------------------------
            for (String testcaseId : testcaseMap.keySet()) {

                System.out.println("\n==============================");
                System.out.println("🚀 Executing Testcase: " + testcaseId);
                System.out.println("==============================");

                // Initialize testcase document
                TestcaseReporter.addStep(driver, testcaseId,
                        "After LPN ASN Creation");


                try {
                    // -------------------------------------------------
                    // STEP 1: Login
                    // -------------------------------------------------
                    driver = login(testcaseId);

                    // -------------------------------------------------
                    // STEP 2: Open ASN module
                    // -------------------------------------------------
                    openASNModule(testcaseId);

                    // -------------------------------------------------
                    // STEP 3: Process each ASN for this testcase
                    // -------------------------------------------------
                    for (Row row : testcaseMap.get(testcaseId)) {
                        String asn = getCellValue(row.getCell(1)); // ASN column
                        if (asn.isEmpty()) continue;

                        processASN(asn, testcaseId);
                    }

                } finally {
                    // -------------------------------------------------
                    // STEP 4: Cleanup
                    // -------------------------------------------------
                    if (driver != null) {
                        driver.quit();
                    }
                    TestcaseReporter.closeTestcase(testcaseId);
                }
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // =====================================================
    // LOGIN
    // =====================================================

    /**
     * Login using Login.xlsx
     * 📸 Screenshot after:
     *  - Login page load
     *  - Successful login
     */
    private WebDriver login(String testcaseId) throws Exception {

        WebDriverManager.chromedriver().setup();
        WebDriver driver = new ChromeDriver();
        driver.manage().window().maximize();

        Map<String, String> login = readLogin();

        driver.get(login.get("LOGIN_URL"));

        // 📸 Screenshot: Login page opened
        TestcaseReporter.addStep(driver, testcaseId, "Opened Login page");

        driver.findElement(By.id("username")).sendKeys(login.get("USERNAME"));
        driver.findElement(By.id("password")).sendKeys(login.get("PASSWORD"));
        driver.findElement(By.id("kc-login")).click();

        new WebDriverWait(driver, Duration.ofSeconds(30))
                .until(ExpectedConditions.presenceOfElementLocated(By.tagName("ion-app")));

        // 📸 Screenshot: Login successful
        TestcaseReporter.addStep(driver, testcaseId, "Login successful");

        return driver;
    }

    // =====================================================
    // NAVIGATION
    // =====================================================

    /**
     * Open ASN module
     * 📸 Screenshot after menu open and ASN open
     */
    private void openASNModule(String testcaseId) {

        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(20));

        WebElement menuToggle = wait.until(
                ExpectedConditions.elementToBeClickable(
                        By.cssSelector("ion-button[data-component-id='menu-toggle-button']"))
        );
        menuToggle.click();

        // 📸 Screenshot: Menu opened
        TestcaseReporter.addStep(driver, testcaseId, "Menu opened");

        WebElement searchBox = wait.until(
                ExpectedConditions.presenceOfElementLocated(
                        By.xpath("//input[@placeholder='Search Menu...']"))
        );
        searchBox.sendKeys("ASNS");

        WebElement asnBtn = wait.until(ExpectedConditions.elementToBeClickable(By.id("ASN")));
        asnBtn.click();

        // 📸 Screenshot: ASN module opened
        TestcaseReporter.addStep(driver, testcaseId, "ASN module opened");
    }

    // =====================================================
    // ITEM-LEVEL VALIDATION
    // =====================================================

    /**
     * Process one ASN:
     * - Search ASN
     * - Open ASN card
     * - Validate received quantity
     * - Navigate to LPN (if qty > 0)
     *
     * 📸 Screenshot after every sub-step
     */
    private void processASN(String asn, String testcaseId) throws Exception {

        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(20));

        WebElement searchInput = wait.until(
                ExpectedConditions.elementToBeClickable(By.xpath("//input[contains(@id,'ion-input')]"))
        );
        searchInput.clear();
        searchInput.sendKeys(asn);
        searchInput.sendKeys(Keys.ENTER);

        // 📸 Screenshot: ASN searched
        TestcaseReporter.addStep(driver, testcaseId, "ASN searched: " + asn);

        WebElement asnCard = wait.until(
                ExpectedConditions.elementToBeClickable(
                        By.cssSelector("card-view[data-component-id='Card-View'] div.card-row.primary"))
        );
        asnCard.click();

        // 📸 Screenshot: ASN card opened
        TestcaseReporter.addStep(driver, testcaseId, "ASN card opened: " + asn);

        WebElement receivedQtyElement = wait.until(
                ExpectedConditions.presenceOfElementLocated(
                        By.cssSelector("span[data-component-id='TotalReceivedQuantity']"))
        );

        String qtyText = receivedQtyElement.getText().trim();
        int receivedQty = 0;
        try {
            receivedQty = Integer.parseInt(qtyText);
        } catch (Exception ignored) {}

        // 📸 Screenshot: Received quantity validation
        TestcaseReporter.addStep(
                driver,
                testcaseId,
                "Received Quantity for ASN " + asn + " = " + receivedQty
        );

        if (receivedQty > 0) {
            WebElement relatedLinksBtn = wait.until(
                    ExpectedConditions.elementToBeClickable(
                            By.cssSelector("button[data-component-id='relatedLinks']"))
            );
            relatedLinksBtn.click();

            // 📸 Screenshot: Related Links opened
            TestcaseReporter.addStep(driver, testcaseId, "Related Links opened");

            WebElement lpnBtn = wait.until(
                    ExpectedConditions.elementToBeClickable(
                            By.cssSelector("ion-item[data-component-id='LPN']"))
            );
            lpnBtn.click();

            // 📸 Screenshot: Navigated to LPN screen
            TestcaseReporter.addStep(driver, testcaseId, "Navigated to LPN screen");
        }
    }

    // =====================================================
    // EXCEL HELPERS
    // =====================================================

    /**
     * Groups rows by Testcase (TST_*)
     */
    private Map<String, List<Row>> readTestcasesFromExcel() throws Exception {

        Map<String, List<Row>> map = new LinkedHashMap<>();

        try (FileInputStream fis = new FileInputStream(DATA_EXCEL_PATH);
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet sheet = wb.getSheet("ItemlvlValidation");
            if (sheet == null)
                throw new RuntimeException("Sheet 'ItemlvlValidation' not found");

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
     * Read login details from Login.xlsx
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
