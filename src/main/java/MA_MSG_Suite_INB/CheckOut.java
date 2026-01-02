package MA_MSG_Suite_INB;

import MA_MSG_Suite_OB.DocPathManager;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.FileInputStream;
import java.time.Duration;
import java.util.*;

public class CheckOut {

    private WebDriver driver;
    public static String docPathLocal;

    // ✅ Central Excel path (aligned with Checkin)
    private static final String DATA_EXCEL_PATH = ExcelReaderIB.DATA_EXCEL_PATH;

    public CheckOut(WebDriver driver) {
        this.driver = driver;
    }

    // ==========================
    // DATA HOLDER
    // ==========================
    static class CheckOutData {
        String testcase;
        String trailerId;
    }

    // ==========================
    // READ EXCEL & GROUP BY TESTCASE
    // ==========================
    private static Map<String, List<CheckOutData>> readCheckoutExcel() throws Exception {

        Map<String, List<CheckOutData>> grouped = new LinkedHashMap<>();

        try (FileInputStream fis = new FileInputStream(DATA_EXCEL_PATH);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheet("Checkin");
            if (sheet == null) {
                throw new RuntimeException("Sheet 'Checkin' not found in Excel file!");
            }

            Row header = sheet.getRow(0);
            Map<String, Integer> cols = new HashMap<>();

            for (int i = 0; i < header.getLastCellNum(); i++) {
                cols.put(header.getCell(i).getStringCellValue().trim(), i);
            }

            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;

                String tc = row.getCell(cols.get("Testcase"),
                        Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString().trim();

                if (!tc.matches("TST_\\d+")) continue;

                String trailerId =
                        row.getCell(cols.get("Trailer"),
                                Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString().trim();

                if (trailerId.isEmpty()) continue;

                CheckOutData data = new CheckOutData();
                data.testcase = tc;
                data.trailerId = trailerId;

                grouped.computeIfAbsent(tc, k -> new ArrayList<>()).add(data);
            }
        }
        return grouped;
    }

    // ==========================
    // EXECUTION (CALLED FROM CONTROLLER)
    // ==========================
    public void execute() throws Exception {

        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(15));
        JavascriptExecutor js = (JavascriptExecutor) driver;

        Map<String, List<CheckOutData>> testcases = readCheckoutExcel();

        for (Map.Entry<String, List<CheckOutData>> entry : testcases.entrySet()) {

            String testcase = entry.getKey();
            List<CheckOutData> rows = entry.getValue();

            docPathLocal = IBDocPathManager.getOrCreateDocPath(
                    ExcelReaderIB.DOC_FILEPATH, testcase);

            System.out.println("▶ Executing Checkout Testcase: " + testcase);

            // ==========================
            // Navigate to Check-Out
            // ==========================
            WebElement menuToggle = wait.until(
                    ExpectedConditions.elementToBeClickable(
                            By.cssSelector("ion-button[data-component-id='menu-toggle-button']")
                    )
            );
            menuToggle.click();

            WebElement searchBox = wait.until(
                    ExpectedConditions.presenceOfElementLocated(
                            By.xpath("//input[@placeholder='Search Menu...']")
                    )
            );
            searchBox.sendKeys("Check Out");

            WebElement checkOutBtn = wait.until(
                    ExpectedConditions.elementToBeClickable(
                            By.xpath("//button[@data-component-id='CheckOut']")
                    )
            );
            checkOutBtn.click();

            IBDocPathManager.captureScreenshot(driver, "Menu_CheckOut");
            IBDocPathManager.saveSharedDocument();

            Thread.sleep(5000);

            WebElement filterBtnHost = wait.until(
                    ExpectedConditions.presenceOfElementLocated(
                            By.xpath("(//ion-button[contains(@class,'toggle-button')])[3]")
                    )
            );

            js.executeScript(
                    "arguments[0].shadowRoot.querySelector('.button-inner').click();",
                    filterBtnHost
            );

            WebElement dropdownHost = driver.findElement(
                    By.cssSelector("ion-button[data-component-id='CheckOutTrailer-Trailer-chevron-down']")
            );

            js.executeScript(
                    "arguments[0].shadowRoot.querySelector('button.button-native').click();",
                    dropdownHost
            );

            IBDocPathManager.captureScreenshot(driver, "Filter_Open");
            IBDocPathManager.saveSharedDocument();

            // ==========================
            // ITERATE ALL ROWS PER TESTCASE
            // ==========================
            for (CheckOutData data : rows) {

                WebElement trailerIdInputField = wait.until(
                        ExpectedConditions.elementToBeClickable(
                                By.xpath("//ion-input[@data-component-id='TrailerId-lookup-dialog-filter-input']//input")
                        )
                );

                trailerIdInputField.click();
                trailerIdInputField.clear();
                trailerIdInputField.sendKeys(data.trailerId, Keys.ENTER);

                IBDocPathManager.captureScreenshot(driver, "Enter_TrailerId_" + data.trailerId);
                IBDocPathManager.saveSharedDocument();

                try {
                    WebElement cardPanelRow = wait.until(
                            ExpectedConditions.elementToBeClickable(
                                    By.cssSelector(
                                            "card-panel card-view[data-component-id='Card-View'] " +
                                                    ".card-row.primary[tabindex='0']"
                                    )
                            )
                    );

                    js.executeScript("arguments[0].scrollIntoView(true);", cardPanelRow);
                    cardPanelRow.click();
                } catch (StaleElementReferenceException e) {
                    WebElement retry = wait.until(
                            ExpectedConditions.elementToBeClickable(
                                    By.cssSelector(
                                            "card-panel card-view[data-component-id='Card-View'] " +
                                                    ".card-row.primary[tabindex='0']"
                                    )
                            )
                    );
                    js.executeScript("arguments[0].scrollIntoView(true);", retry);
                    retry.click();
                }

                IBDocPathManager.captureScreenshot(driver, "Select_Card");
                IBDocPathManager.saveSharedDocument();

                WebElement checkoutHost = driver.findElement(
                        By.cssSelector("ion-button[data-component-id='footer-panel-action-CheckOut']")
                );

                js.executeScript("arguments[0].scrollIntoView(true);", checkoutHost);
                js.executeScript(
                        "arguments[0].shadowRoot.querySelector('button.button-native').click();",
                        checkoutHost
                );

                IBDocPathManager.captureScreenshot(driver, "Click_CheckOut");
                IBDocPathManager.saveSharedDocument();

                WebElement saveHost = driver.findElement(
                        By.cssSelector("ion-button[data-component-id='save-btn']")
                );

                js.executeScript("arguments[0].scrollIntoView(true);", saveHost);
                js.executeScript(
                        "arguments[0].shadowRoot.querySelector('button.button-native').click();",
                        saveHost
                );

                IBDocPathManager.captureScreenshot(driver, "Save_CheckOut");
                IBDocPathManager.saveSharedDocument();
            }
        }
    }
}
