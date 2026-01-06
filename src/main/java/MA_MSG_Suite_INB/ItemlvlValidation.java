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

    public void execute(WebDriver driver) {
        try {
            Map<String, List<String>> testcaseAsns = readAsnsByTestcase();
            if (testcaseAsns.isEmpty()) {
                System.out.println("⚠ No Item ASNs found in workbook — skipping Item validation.");
                return;
            }

            for (Map.Entry<String, List<String>> entry : testcaseAsns.entrySet()) {
                String testcaseName = entry.getKey();
                List<String> asnList = entry.getValue();

                System.out.println("\n===== Item-level validation for Testcase: " + testcaseName + " =====");
                System.out.println("🔎 ASNs: " + asnList);

                docPathLocal = IBDocPathManager.getOrCreateDocPath(ExcelReaderIB.DOC_FILEPATH, testcaseName);
                System.out.println("📂 Screenshot doc path: " + docPathLocal);

                validateItemForAsn(driver, testcaseName, asnList);

                System.out.println("===== Finished Item validation for Testcase: " + testcaseName + " =====\n");
            }
        } catch (Exception e) {
            System.out.println("❌ Item validation failed with error: " + e.getMessage());
            e.printStackTrace();
        }
    }

    //ASN validation
    //--------------------------------------
    //ASN validation


    private void validateItemForAsn(WebDriver driver, String testcaseName, List<String> asnList) {
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(30));
        JavascriptExecutor js = (JavascriptExecutor) driver;

        try {
            Thread.sleep(6000);

            // Menu toggle step 1
            WebElement menuToggle = wait.until(ExpectedConditions.elementToBeClickable(
                    By.cssSelector("ion-button[data-component-id='menu-toggle-button']")));
            try { menuToggle.click(); }
            catch (Exception e) { js.executeScript("arguments[0].click();", menuToggle); }

            // Search menu step 2
            WebElement searchBox = wait.until(ExpectedConditions.presenceOfElementLocated(
                    By.xpath("//input[@placeholder='Search Menu...']")));
            searchBox.clear();
            searchBox.sendKeys("ASNs");

            WebElement asnsButton = wait.until(ExpectedConditions.elementToBeClickable(
                    By.xpath("//button[@data-component-id='ASNs']")));
            try { asnsButton.click(); }
            catch (Exception e) { js.executeScript("arguments[0].click();", asnsButton); }

            Thread.sleep(3000);



// Adaptive block: try input field first, else fall back to steps 3 & 4
            WebElement asnInputField = null;
            try {
                // Step 5 directly
                WebElement asnInputHost = wait.until(ExpectedConditions.presenceOfElementLocated(
                        By.cssSelector("ion-input[data-component-id='AsnId']")));
                asnInputField = asnInputHost.findElement(By.cssSelector("input.native-input"));

                wait.until(ExpectedConditions.visibilityOf(asnInputField));
                js.executeScript("arguments[0].scrollIntoView({block:'center'});", asnInputField);
                js.executeScript("arguments[0].click();", asnInputField);

                System.out.println("✅ ASN input field visible, skipping steps 3 & 4.");
            } catch (TimeoutException e) {
                System.out.println("⚠ ASN input field not visible, performing steps 3 & 4 first.");

                // Step 3: Expand filter
                WebElement filterBtnHost = wait.until(ExpectedConditions.elementToBeClickable(
                        By.xpath("(//ion-button[contains(@class,'toggle-button')])[3]")));
                js.executeScript("arguments[0].shadowRoot.querySelector('.button-inner').click();", filterBtnHost);
                Thread.sleep(3000);

                // Step 4: Dropdown chevron
                try {
                    WebElement dropdownHost = wait.until(ExpectedConditions.presenceOfElementLocated(
                            By.cssSelector("ion-button[data-component-id='ASN-ASN-chevron-down']")));
                    WebElement dropdownButton = (WebElement) js.executeScript(
                            "return arguments[0].shadowRoot.querySelector('button.button-native')", dropdownHost);
                    wait.until(ExpectedConditions.elementToBeClickable(dropdownButton)).click();
                } catch (TimeoutException ignored) {
                    System.out.println("⚠ ASN chevron not found, continuing anyway.");
                }

                Thread.sleep(3000);

                // Now Step 5: Input field
                WebElement asnInputHost = wait.until(ExpectedConditions.presenceOfElementLocated(
                        By.cssSelector("ion-input[data-component-id='AsnId']")));
                asnInputField = asnInputHost.findElement(By.cssSelector("input.native-input"));

                wait.until(ExpectedConditions.visibilityOf(asnInputField));
                js.executeScript("arguments[0].scrollIntoView({block:'center'});", asnInputField);
                js.executeScript("arguments[0].click();", asnInputField);
            }



            // Enter each ASN one by one step 6
            for (String asnValue : asnList) {
                if (asnValue != null && !asnValue.isEmpty()) {
                    // Try Clear button first
                    try {
                        WebElement clearBtn = wait.until(ExpectedConditions.presenceOfElementLocated(
                                By.cssSelector("ion-button[data-component-id='Clear']")));
                        js.executeScript("arguments[0].click();", clearBtn);
                        Thread.sleep(1000);
                    } catch (TimeoutException ignored) {
                        js.executeScript("arguments[0].value = '';", asnInputField);
                        js.executeScript("arguments[0].dispatchEvent(new Event('input'));", asnInputField);
                    }

                    // Enter ASN value step 7
                    js.executeScript(
                            "arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('input'));",
                            asnInputField, asnValue);

                    System.out.println("Entered ASN from Excel: " + asnValue);

                    asnInputField.sendKeys(Keys.ENTER);
                    Thread.sleep(2000);
                }
            }

            // ASN card step 8
            WebElement asnCard = wait.until(ExpectedConditions.elementToBeClickable(
                    By.cssSelector("card-view[data-component-id='Card-View'] div.card-row.primary")));
            asnCard.click();
            Thread.sleep(3000);

            IBDocPathManager.captureScreenshot(driver, "Created ASN");
            IBDocPathManager.saveSharedDocument();
            System.out.println("✅ ASN screenshot captured for " + testcaseName);

            // Received qty step 9
            WebElement receivedQty = wait.until(ExpectedConditions.presenceOfElementLocated(
                    By.cssSelector("span[data-component-id='TotalReceivedQuantity']")));
            int qty = Integer.parseInt(receivedQty.getText().trim());
            System.out.println(testcaseName + " → Received Qty = " + qty);

            if (qty > 0) {
                WebElement relatedLinks = wait.until(ExpectedConditions.elementToBeClickable(
                        By.cssSelector("button[data-component-id='relatedLinks']")));
                relatedLinks.click();

                WebElement lpnBtn = wait.until(ExpectedConditions.elementToBeClickable(
                        By.cssSelector("ion-item[data-component-id='LPN']")));
                lpnBtn.click();
                Thread.sleep(3000);

                IBDocPathManager.captureScreenshot(driver, "Created LPN");
                IBDocPathManager.saveSharedDocument();
                System.out.println("✅ LPN screenshot captured for " + testcaseName);
            }

        } catch (Exception e) {
            System.out.println("❌ Item validation failed for testcase " + testcaseName + ": " + e.getMessage());
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
            if (sheet == null) throw new RuntimeException("Sheet 'Item_ASN' not found");

            Row header = sheet.getRow(0);
            int asnCol = -1, tcCol = -1;
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

                String[] splitAsns = asn.split(",");
                for (String a : splitAsns) {
                    if (!a.trim().isEmpty()) {
                        map.computeIfAbsent(tc, k -> new ArrayList<>()).add(a.trim());
                    }
                }
            }
        }
        return map;
    }

    private String getCellValue(Cell cell) {
        return cell == null ? "" : new DataFormatter().formatCellValue(cell).trim();
    }
}
