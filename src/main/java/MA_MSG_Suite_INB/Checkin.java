package MA_MSG_Suite_INB;

import MA_MSG_Suite_OB.DocPathManager;
import org.openqa.selenium.*;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.time.Duration;
import java.util.*;

public class Checkin {

    private WebDriver driver;
    public static String docPathLocal;

    private static final String DATA_EXCEL_PATH = ExcelReaderIB.DATA_EXCEL_PATH;

    public Checkin(WebDriver driver) {
        this.driver = driver;
    }

    // ==========================
    // DATA HOLDER
    // ==========================
    static class CheckinData {
        String testcase;
        String asn;
        String carrier;
        String trailer;
        String trailerType;
        String dockDoor;
    }

    // ==========================
    // READ EXCEL & GROUP BY TESTCASE
    // ==========================
    private static Map<String, List<CheckinData>> readCheckinExcel() throws Exception {

        Map<String, List<CheckinData>> grouped = new LinkedHashMap<>();

        try (FileInputStream fis = new FileInputStream(DATA_EXCEL_PATH);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheet("Checkin");
            if (sheet == null) {
                throw new RuntimeException("Sheet 'Checkin' not found!");
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

                CheckinData data = new CheckinData();
                data.testcase = tc;
                data.asn = row.getCell(cols.get("ASNs")).toString().trim();
                data.carrier = row.getCell(cols.get("Carrier")).toString().trim();
                data.trailer = row.getCell(cols.get("Trailer")).toString().trim();
                data.trailerType = row.getCell(cols.get("Trailer type")).toString().trim();
                data.dockDoor = row.getCell(cols.get("Dock Door")).toString().trim();

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
        Map<String, List<CheckinData>> testcases = readCheckinExcel();

        for (Map.Entry<String, List<CheckinData>> entry : testcases.entrySet()) {

            String testcase = entry.getKey();
            List<CheckinData> rows = entry.getValue();

            docPathLocal = IBDocPathManager.getOrCreateDocPath(
                    ExcelReaderIB.DOC_FILEPATH, testcase);

            System.out.println("▶ Executing Checkin Testcase: " + testcase);

            for (CheckinData data : rows) {

                // ==========================
                // Menu
                // ==========================
                Thread.sleep(6000);
                wait.until(ExpectedConditions.elementToBeClickable(
                        By.cssSelector("ion-button[data-component-id='menu-toggle-button']"))).click();

                wait.until(ExpectedConditions.presenceOfElementLocated(
                        By.xpath("//input[@placeholder='Search Menu...']"))).sendKeys("Check In");

                wait.until(ExpectedConditions.elementToBeClickable(
                        By.xpath("//button[@data-component-id='CheckIn']"))).click();

                wait.until(ExpectedConditions.presenceOfElementLocated(
                        By.xpath("//span[@data-component-id='Check-InInfo']"))).click();

                IBDocPathManager.captureScreenshot(driver, "Menu_Navigation");
                IBDocPathManager.saveSharedDocument();

                // Step 1 Carrier
                WebElement carrierInput = wait.until(ExpectedConditions.elementToBeClickable(
                        By.cssSelector("ion-input[data-component-id='CarrierId'] input.native-input")));
                carrierInput.clear();
                carrierInput.sendKeys(data.carrier);
                IBDocPathManager.captureScreenshot(driver, "Step1_Carrier");
                IBDocPathManager.saveSharedDocument();

                // Step 2 Visit Type
                wait.until(ExpectedConditions.elementToBeClickable(
                        By.cssSelector("ion-input[data-component-id='VisitType'] input.native-input"))).click();
                wait.until(ExpectedConditions.elementToBeClickable(
                        By.cssSelector("button[data-component-id='LiveUnLoad-dropdown-option']"))).click();
                IBDocPathManager.captureScreenshot(driver, "Step2_VisitType");
                IBDocPathManager.saveSharedDocument();

                // Step 3 Dock Door
                // Step 1: Locate ion-input wrapper (NOT the native input)
                WebElement dockDoorIonInput = wait.until(
                        ExpectedConditions.presenceOfElementLocated(
                                By.cssSelector("ion-input[data-component-id='LocationId']")
                        )
                );

                JavascriptExecutor js = (JavascriptExecutor) driver;
                js.executeScript("arguments[0].scrollIntoView({block:'center'});", dockDoorIonInput);
                js.executeScript("arguments[0].click();", dockDoorIonInput);

                WebElement dockDoorNativeInput = dockDoorIonInput.findElement(
                        By.cssSelector("input.native-input")
                );

                js.executeScript("arguments[0].focus();", dockDoorNativeInput);
                dockDoorNativeInput.clear();
                dockDoorNativeInput.sendKeys(data.dockDoor);

                IBDocPathManager.captureScreenshot(driver, "Step3_DockDoor");
                IBDocPathManager.saveSharedDocument();

                // Step 4 TrailerType
                WebElement trailertypeInput = wait.until(ExpectedConditions.elementToBeClickable(
                        By.cssSelector("ion-input[data-component-id='EquipmentTypeId'] input.native-input")));
                trailertypeInput.click();
                trailertypeInput.clear();
                trailertypeInput.sendKeys(data.trailerType);

                IBDocPathManager.captureScreenshot(driver, "Step4_TrailerType");
                IBDocPathManager.saveSharedDocument();

                // Step 5 Mandatory TrailerId
                WebElement mandatoryTrailerInput = wait.until(
                        ExpectedConditions.presenceOfElementLocated(
                                By.xpath("//span[@data-component-id='TrailerId-label' and contains(@class,'has-required-field')]" +
                                        "/ancestor::widget-label/following-sibling::ion-row" +
                                        "//ion-input[@data-component-id='TrailerId']//input")
                        )
                );

                ((JavascriptExecutor) driver).executeScript("arguments[0].focus();", mandatoryTrailerInput);
                mandatoryTrailerInput.sendKeys(data.trailer);

                IBDocPathManager.captureScreenshot(driver, "Step5_TrailerId");
                IBDocPathManager.saveSharedDocument();

                // Step 6 Enter ASN
                // Expand Trailer Content Details (unchanged)
                WebElement trailerExpand = driver.findElement(
                        By.cssSelector("div.expand-header-container[data-component-id*='TrailerContentDetails']")
                );
                ((JavascriptExecutor) driver).executeScript(
                        "arguments[0].scrollIntoView({block: 'center'});", trailerExpand);
                ((JavascriptExecutor) driver).executeScript("arguments[0].click();", trailerExpand);

                // Locate ion-input wrapper (NOT native input)
                WebElement asnIonInput = wait.until(
                        ExpectedConditions.presenceOfElementLocated(
                                By.cssSelector("ion-input[data-component-id='Asn']")
                        )
                );

                js.executeScript("arguments[0].scrollIntoView({block:'center'});", asnIonInput);
                js.executeScript("arguments[0].click();", asnIonInput);

                WebElement asnNativeInput = asnIonInput.findElement(
                        By.cssSelector("input.native-input")
                );

                js.executeScript("arguments[0].focus();", asnNativeInput);
                asnNativeInput.clear();
                asnNativeInput.sendKeys(data.asn);

                IBDocPathManager.captureScreenshot(driver, "Step6_ASN");
                IBDocPathManager.saveSharedDocument();

                //Step 7 Submit Check in
                WebElement submitCheckin = wait.until(
                        ExpectedConditions.presenceOfElementLocated(
                                By.xpath("//ion-button[@data-component-id='checkin-btn']")
                        )
                );

                ((JavascriptExecutor) driver).executeScript("arguments[0].click();", submitCheckin);

                IBDocPathManager.captureScreenshot(driver, "Step7_Submit");
                IBDocPathManager.saveSharedDocument();
            }
        }
    }
}
