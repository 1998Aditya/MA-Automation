package MA_MSG_Suite_INB;

import org.apache.poi.ss.usermodel.*;
import org.openqa.selenium.*;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.FileInputStream;
import java.time.Duration;
import java.util.*;

/**
 * =========================================================
 * ItemlvlValidation
 * ---------------------------------------------------------
 * Purpose:
 * - Validate Item-level receiving for ASN
 * - Uses ASN IDs from Item_ASN sheet
 * - Executes for all testcases found in Excel
 * =========================================================
 */
public class ItemlvlValidation {
    public static String docPathLocal;

    private static final String DATA_EXCEL_PATH = ExcelReaderIB.DATA_EXCEL_PATH;

    /**
     * Entry point (called from controller with driver only)
     */
    public void execute(WebDriver driver) {
        try {
            Map<String, List<String>> testcaseAsns = readAsnsByTestcase();

            if (testcaseAsns.isEmpty()) {
                System.out.println("⚠ No Item ASNs found in workbook — skipping Item validation.");
                return;
            }

            for (Map.Entry<String, List<String>> entry : testcaseAsns.entrySet()) {
                String testcase = entry.getKey();
                List<String> asns = entry.getValue();

                System.out.println("\n===== Item-level validation for Testcase: " + testcase + " =====");
                System.out.println("🔎 ASNs: " + asns);

                docPathLocal = IBDocPathManager.getOrCreateDocPath(ExcelReaderIB.DOC_FILEPATH, testcase);
                System.out.println("📂 Screenshot doc path: " + docPathLocal);

                for (String asn : asns) {
                    validateItemForAsn(driver, asn);
                }

                System.out.println("===== Finished Item validation for Testcase: " + testcase + " =====\n");
            }

        } catch (Exception e) {
            System.out.println("❌ Item validation failed with error: " + e.getMessage());
            e.printStackTrace();
        }
    }

    // =====================================================
    // ITEM-LEVEL UI VALIDATION
    // =====================================================
    private void validateItemForAsn(WebDriver driver, String asn) {
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(30));

// Pause for 5 seconds
        try {
            Thread.sleep(5000);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }

        try {
            Thread.sleep(6000);
            // Click menu toggle
            WebElement menuToggle = wait.until(
                    ExpectedConditions.elementToBeClickable(
                            By.cssSelector("ion-button[data-component-id='menu-toggle-button']"))
            );
            menuToggle.click();

            // Search menu for "Check In"
            WebElement searchBox = wait.until(
                    ExpectedConditions.presenceOfElementLocated(
                            By.xpath("//input[@placeholder='Search Menu...']"))
            );
            searchBox.sendKeys("ASNs");

           WebElement asnsButton = wait.until(
                    ExpectedConditions.presenceOfElementLocated(
                            By.xpath("//button[@data-component-id='ASNs']"))
            );
            asnsButton.click();

            Thread.sleep(5000);

            // Expand filter section
            JavascriptExecutor js = (JavascriptExecutor) driver;
            WebElement filterBtnHost = wait.until(ExpectedConditions.elementToBeClickable(
                    By.xpath("(//ion-button[contains(@class,'toggle-button')])[3]")));
            js.executeScript("arguments[0].shadowRoot.querySelector('.button-inner').click();", filterBtnHost);

            Thread.sleep(3000);

            // Dropdown expand
            WebElement dropdownHost = wait.until(ExpectedConditions.presenceOfElementLocated(
                    By.cssSelector("ion-button[data-component-id='ASN-ASN-chevron-down']")
            ));
            WebElement dropdownButton = (WebElement) js.executeScript(
                    "return arguments[0].shadowRoot.querySelector('button.button-native')",
                    dropdownHost
            );
            js.executeScript("arguments[0].scrollIntoView(true);", dropdownButton);
            wait.until(ExpectedConditions.elementToBeClickable(dropdownButton)).click();
            System.out.println("ASN chevron-down dropdown clicked successfully.");

            Thread.sleep(3000);

// After dropdown expand
            WebDriverWait wait1= new WebDriverWait(driver, Duration.ofSeconds(20));
            JavascriptExecutor js1 = (JavascriptExecutor) driver;

// Anchor on ion-input host
            WebElement asnInputHost = wait1.until(
                    ExpectedConditions.presenceOfElementLocated(By.cssSelector("ion-input[data-component-id='AsnId']"))
            );

// Find the native input inside
            WebElement asnInputField = asnInputHost.findElement(By.cssSelector("input.native-input"));

// Ensure visible
            wait.until(ExpectedConditions.visibilityOf(asnInputField));

// Scroll and click
            js1.executeScript("arguments[0].scrollIntoView({block:'center'});", asnInputField);
            js1.executeScript("arguments[0].click();", asnInputField);

            System.out.println("ASN input field clicked successfully.");


// Read all ASNs grouped by testcase
            Map<String, List<String>> asnsByTestcase = readAsnsByTestcase();

// Iterate through each testcase and its ASN list
            for (Map.Entry<String, List<String>> entry : asnsByTestcase.entrySet()) {
                String testcaseName = entry.getKey();
                List<String> asnList = entry.getValue();

                System.out.println("Processing testcase: " + testcaseName);

                for (String asnValue : asnList) {
                    if (asnValue != null && !asnValue.isEmpty()) {
                        // Clear field before entering
                        asnInputField.clear();

                        // Use JS to set value and trigger Angular/Ionic binding
                        js.executeScript(
                                "arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('input'));",
                                asnInputField, asnValue
                        );

                        System.out.println("Entered ASN from Excel: " + asnValue);

                        // Press ENTER to confirm
                        asnInputField.sendKeys(Keys.ENTER);

                        Thread.sleep(2000); // wait between entries
                    }
                }
            }











            // Click ASN card
            WebElement asnCard = wait.until(
                    ExpectedConditions.elementToBeClickable(
                            By.cssSelector("card-view[data-component-id='Card-View'] div.card-row.primary"))
            );
            asnCard.click();
            Thread.sleep(3000);

            IBDocPathManager.captureScreenshot(driver, "Created ASN");
            IBDocPathManager.saveSharedDocument();
            System.out.println("✅ ASN " + asn + " screenshot captured");

            // Check received quantity
            WebElement receivedQty = wait.until(
                    ExpectedConditions.presenceOfElementLocated(
                            By.cssSelector("span[data-component-id='TotalReceivedQuantity']"))
            );

            int qty = Integer.parseInt(receivedQty.getText().trim());
            System.out.println("ASN " + asn + " → Received Qty = " + qty);

            if (qty > 0) {
                WebElement relatedLinks = wait.until(
                        ExpectedConditions.elementToBeClickable(
                                By.cssSelector("button[data-component-id='relatedLinks']"))
                );
                relatedLinks.click();

                WebElement lpnBtn = wait.until(
                        ExpectedConditions.elementToBeClickable(
                                By.cssSelector("ion-item[data-component-id='LPN']"))
                );
                lpnBtn.click();
                Thread.sleep(3000);

                IBDocPathManager.captureScreenshot(driver, "Created LPN");
                IBDocPathManager.saveSharedDocument();
                System.out.println("✅ LPN screenshot captured for ASN " + asn);
            }

        } catch (Exception e) {
            System.out.println("❌ Item validation failed for ASN " + asn + ": " + e.getMessage());
            e.printStackTrace();
        }
    }

    // =====================================================
    // EXCEL READER
    // =====================================================
    private Map<String, List<String>> readAsnsByTestcase() throws Exception {
        Map<String, List<String>> map = new LinkedHashMap<>();

        try (FileInputStream fis = new FileInputStream(DATA_EXCEL_PATH);
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet sheet = wb.getSheet("Item_ASN");
            if (sheet == null) {
                throw new RuntimeException("Sheet 'Item_ASN' not found in " + DATA_EXCEL_PATH);
            }

            Row header = sheet.getRow(0);
            int asnCol = -1;
            int tcCol = -1;

            for (int i = 0; i < header.getLastCellNum(); i++) {
                String h = header.getCell(i).getStringCellValue();
                if ("AsnId".equalsIgnoreCase(h)) asnCol = i;
                if ("Testcase".equalsIgnoreCase(h)) tcCol = i;
            }

            Iterator<Row> it = sheet.iterator();
            it.next(); // skip header

            while (it.hasNext()) {
                Row r = it.next();
                String tc = getCellValue(r.getCell(tcCol));
                String asn = getCellValue(r.getCell(asnCol));
                if (!tc.matches("TST_\\d+") || asn.isEmpty()) continue;
                map.computeIfAbsent(tc, k -> new ArrayList<>()).add(asn);
            }
        }
        return map;
    }

    private String getCellValue(Cell cell) {
        return cell == null ? "" : new DataFormatter().formatCellValue(cell).trim();
    }
}

