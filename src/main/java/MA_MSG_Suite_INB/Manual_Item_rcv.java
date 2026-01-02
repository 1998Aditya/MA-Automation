package MA_MSG_Suite_INB;

import MA_MSG_Suite_OB.DocPathManager;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.FileInputStream;
import java.time.Duration;
import java.util.List;
import java.util.Set;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Manual_Item_rcv {

    private WebDriver driver;
    public static String docPathLocal ;
    // ✅ Excel data path now comes from central ExcelReaderIB
    private static final String DATA_EXCEL_PATH = ExcelReaderIB.DATA_EXCEL_PATH;

    public Manual_Item_rcv(WebDriver driver) {
        this.driver = driver;
    }

    public void execute() throws InterruptedException {
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(15));

        // Step 1: Click menu toggle
        WebElement menuToggle = wait.until(
                ExpectedConditions.elementToBeClickable(
                        By.cssSelector("ion-button[data-component-id='menu-toggle-button']"))
        );
        menuToggle.click();

        // Step 2: Search menu and click WM Mobile
        WebElement searchBox = wait.until(
                ExpectedConditions.presenceOfElementLocated(
                        By.xpath("//input[@placeholder='Search Menu...']"))
        );
        searchBox.sendKeys("WM Mobile");

        WebElement WM_Mobile = wait.until(
                ExpectedConditions.presenceOfElementLocated(
                        By.xpath("//button[@data-component-id='WMMobile']"))
        );
        WM_Mobile.click();

        // Step 3: Switch to new tab
        String originalWindow = driver.getWindowHandle();
        String secondTabHandle = "";
        Set<String> windows = driver.getWindowHandles();
        for (String window : windows) {
            if (!window.equals(originalWindow)) {
                driver.switchTo().window(window);
                secondTabHandle = window; // store second tab handle
                break;
            }
        }

        // Step 4: Read values from Excel
        try (FileInputStream fis = new FileInputStream(DATA_EXCEL_PATH);
             Workbook workbook = WorkbookFactory.create(fis)) {

            Sheet sheet = workbook.getSheet("item_rcv");
            if (sheet == null) {
                throw new RuntimeException("Sheet 'item_rcv' not found in Excel file!");
            }

            // Group rows by Testcase (TST_#) and process each testcase in sequence
            Map<String, List<Row>> grouped = groupRows(sheet);
            if (grouped.isEmpty()) {
                // fallback: if no TST_* values found, process entire sheet as before
                System.out.println("⚠ No Testcase groups found — processing whole sheet in Inbound.");
                // simulate a single group containing all data rows
                List<Row> allRows = new ArrayList<>();
                Iterator<Row> itAll = sheet.iterator();
                if (itAll.hasNext()) itAll.next(); // skip header
                while (itAll.hasNext()) {
                    Row r = itAll.next();
                    if (r != null) allRows.add(r);
                }
                grouped.put("ALL_ROWS", allRows);
            }

            // We'll derive column indexes from header (row 0)
            Row headerRow = sheet.getRow(0);
            if (headerRow == null) throw new RuntimeException("Sheet 'item_rcv' has no header row!");

            int trxnCol = -1, dockDoorCol = -1, asnCol = -1, lpnCol = -1, itemCol = -1;
            int enterQtyCol = -1, directToPalletCol = -1, testcaseColIndex = -1;

            for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                String colName = getCellValue(headerRow.getCell(i));
                if (colName.equalsIgnoreCase("Trxn_Name")) {
                    trxnCol = i;
                } else if (colName.equalsIgnoreCase("Dock door")) {
                    dockDoorCol = i;
                } else if (colName.equalsIgnoreCase("ASN")) {
                    asnCol = i;
                } else if (colName.equalsIgnoreCase("LPN")) {
                    lpnCol = i;
                } else if (colName.equalsIgnoreCase("Item")) {
                    itemCol = i;
                } else if (colName.equalsIgnoreCase("Enter_Quantity_UNIT")) {
                    enterQtyCol = i;
                } else if (colName.equalsIgnoreCase("Direct_To_Pallet")) {
                    directToPalletCol = i;
                } else if (colName.equalsIgnoreCase("Testcase")) {
                    testcaseColIndex = i;
                }
            }

            if (trxnCol == -1 || dockDoorCol == -1 || asnCol == -1 ||
                    lpnCol == -1 || itemCol == -1 || enterQtyCol == -1 || directToPalletCol == -1) {
                throw new RuntimeException("One or more required columns (Trxn_Name, Dock door, ASN, LPN, Item, Enter_Quantity_UNIT, Direct_To_Pallet) not found in Excel!");
            }

            // Process each Testcase group in sheet order (LinkedHashMap preserves order)
            for (Map.Entry<String, List<Row>> entry : grouped.entrySet()) {
                String testcase = entry.getKey();
                List<Row> rowsForTestcase = entry.getValue();

                System.out.println("▶ Running Testcase group: " + testcase + " (" + rowsForTestcase.size() + " rows)");
                docPathLocal = IBDocPathManager.getOrCreateDocPath(ExcelReaderIB.DOC_FILEPATH, testcase);//Screenshot
                System.out.println("Path"+docPathLocal); //Screenshot Doc
                // === Step 5: Enter Trxn_Name in search box on new tab ===
                // Use first row of this testcase group to pick transaction and dock door
                Row firstRow = rowsForTestcase.get(0);
                String trxnName = getCellValue(firstRow.getCell(trxnCol));
                String dockDoor = getCellValue(firstRow.getCell(dockDoorCol));

                WebElement newTabSearch = wait.until(
                        ExpectedConditions.presenceOfElementLocated(
                                By.xpath("//input[@placeholder='Search']"))
                );
                Thread.sleep(1000);
                newTabSearch.click();
                newTabSearch.sendKeys(trxnName);
                Thread.sleep(3000);
                IBDocPathManager.captureScreenshot(driver,"ENTER TRANSACTION"); //Screenshot
                IBDocPathManager.saveSharedDocument();                               //Screenshot Doc
                System.out.println("ENTER TRANSACTION screenshot done");
                newTabSearch.sendKeys(Keys.ENTER);

                // Step 6: Match result by text and click it
                List<WebElement> allElements = driver.findElements(By.xpath("//*"));
                for (WebElement el : allElements) {
                    try {
                        if (el.isDisplayed() && el.getText().trim().equals(trxnName)) {
                            el.click();
                            break;
                        }
                    } catch (Exception ignored) {}
                }

                // Step 7: Enter Dock Door
//                WebElement dockDoorInput = wait.until(
//                        ExpectedConditions.presenceOfElementLocated(
//                                By.xpath("//input[@placeholder=\"Scan Dock Door\"]"))
//                );
//                dockDoorInput.click();
//                dockDoorInput.sendKeys(dockDoor);
//                dockDoorInput.sendKeys(Keys.ENTER);

                By locator = By.xpath("//input[@placeholder='Scan Dock Door']");

                WebElement dockDoorInput = wait.until(ExpectedConditions.visibilityOfElementLocated(locator));
                wait.until(ExpectedConditions.elementToBeClickable(locator));
                dockDoorInput.click();               // sometimes needed to focus
                dockDoorInput.clear();
                dockDoorInput.sendKeys(dockDoor);
                Thread.sleep(3000);
                IBDocPathManager.captureScreenshot(driver,"Enter DockDoor"); //Screenshot
                IBDocPathManager.saveSharedDocument();                               //Screenshot Doc
                System.out.println("Enter DockDoor screenshot done");
                dockDoorInput.sendKeys(Keys.ENTER);


                // === Loop: iterate rows for this testcase and group rows by ASN logic preserved ===
                String previousASN = "";
                for (Row row : rowsForTestcase) {
                    if (row == null) continue;

                    String asn = getCellValue(row.getCell(asnCol));
                    String lpn = getCellValue(row.getCell(lpnCol));
                    String item = getCellValue(row.getCell(itemCol));
                    String enterQuantityUnit = getCellValue(row.getCell(enterQtyCol));
                    String directToPallet = getCellValue(row.getCell(directToPalletCol));

                    // If ASN changes, release dock door first
                    if (!asn.equals(previousASN) && !previousASN.isEmpty()) {
                        WebElement releaseDock = wait.until(
                                ExpectedConditions.presenceOfElementLocated(
                                        By.xpath("//button[@data-component-id='action_releasedoor_button']"))
                        );
                        releaseDock.click();
                        Thread.sleep(2000);
                    }

                    // Step 8: Enter ASN if changed
                    if (!asn.equals(previousASN)) {
                        WebElement asnInput;
                        try {
                            asnInput = wait.until(
                                    ExpectedConditions.presenceOfElementLocated(
                                            By.xpath("//input[@placeholder='Asn']"))
                            );
                        } catch (Exception e) {
                            WebElement assocButton = wait.until(
                                    ExpectedConditions.elementToBeClickable(
                                            By.xpath("//button[@data-component-id='action_associateadditionalasn_button']"))
                            );
                            assocButton.click();
                            asnInput = wait.until(
                                    ExpectedConditions.presenceOfElementLocated(
                                            By.xpath("//input[@placeholder='Asn']"))
                            );
                        }
                        Actions actions = new Actions(driver);
                        actions.doubleClick(asnInput).perform();
                        asnInput.sendKeys(Keys.BACK_SPACE);
                        asnInput.sendKeys(asn);
                        Thread.sleep(3000);
                        IBDocPathManager.captureScreenshot(driver,"Enter ASN"); //Screenshot
                        IBDocPathManager.saveSharedDocument();                               //Screenshot Doc
                        System.out.println("Enter ASN screenshot done");
                        asnInput.sendKeys(Keys.ENTER);

                        previousASN = asn;
                    }

                    // Step 9: Enter LPN
                    WebElement lpnInput = wait.until(
                            ExpectedConditions.presenceOfElementLocated(
                                    By.xpath("//input[@placeholder='Scan LPN']"))
                    );
                    lpnInput.click();
                    lpnInput.sendKeys(lpn);
                    Thread.sleep(3000);
                    IBDocPathManager.captureScreenshot(driver,"Enter LPN"); //Screenshot
                    IBDocPathManager.saveSharedDocument();                               //Screenshot Doc
                    System.out.println("Enter LPN screenshot done");
                    lpnInput.sendKeys(Keys.ENTER);
                    Thread.sleep(3000);

                    // Step 10: Switch to first tab and capture item barcode
//                    driver.switchTo().window(originalWindow);
//
//                    WebElement menu = wait.until(
//                            ExpectedConditions.elementToBeClickable(
//                                    By.cssSelector("ion-button[data-component-id='menu-toggle-button']"))
//                    );
//                    menu.click();
//
//                    WebElement searchBox1 = wait.until(
//                            ExpectedConditions.presenceOfElementLocated(
//                                    By.xpath("//input[@placeholder='Search Menu...']"))
//                    );
//                    searchBox1.click();
//                    searchBox1.sendKeys(Keys.chord(Keys.CONTROL, "a"));
//                    searchBox1.sendKeys(Keys.BACK_SPACE);
//                    searchBox1.sendKeys("Items");
//
//                    WebElement Items = wait.until(
//                            ExpectedConditions.presenceOfElementLocated(
//                                    By.xpath("//button[@data-component-id='Items']"))
//                    );
//                    Items.click();
//
//                    WebElement ItemIdFilter = wait.until(
//                            ExpectedConditions.presenceOfElementLocated(
//                                    By.xpath("//ion-input[@data-component-id='ItemId-lookup-dialog-filter-input']//input"))
//                    );
//                    ItemIdFilter.click();
//                    ItemIdFilter.sendKeys(Keys.chord(Keys.CONTROL, "a"));
//                    ItemIdFilter.sendKeys(Keys.BACK_SPACE);
//                    ItemIdFilter.sendKeys(item);
//                    ItemIdFilter.sendKeys(Keys.ENTER);
//
//                    Thread.sleep(5000);
//
//                    String itemBarcode = "";
//                    try {
//                        WebElement primaryBarcodeElement = wait.until(
//                                ExpectedConditions.presenceOfElementLocated(
//                                        By.cssSelector("span[data-component-id='PrimaryBarCode']"))
//                        );
//                        itemBarcode = primaryBarcodeElement.getText().trim();
//                    } catch (Exception e) {
//                        try {
//                            WebElement extendedBarcodeElement = wait.until(
//                                    ExpectedConditions.presenceOfElementLocated(
//                                            By.cssSelector("span[data-component-id='Extended\\.MAUJDSDefaultEANBarcode']"))
//                            );
//                            itemBarcode = extendedBarcodeElement.getText().trim();
//                        } catch (Exception ex) {
//                            System.out.println("❌ No barcode element found on screen!");
//                        }
//                    }
//
//                    driver.switchTo().window(secondTabHandle);
                    String itemBarcode = ItemBarcodeService.getItemBarcode(item);
                    System.out.println("📦 Barcode from API for item " + item + ": " + itemBarcode);

                    // Step 11: Enter Item Barcode
                    WebElement itemInput = wait.until(
                            ExpectedConditions.presenceOfElementLocated(
                                    By.xpath("//input[@placeholder='Scan Item']"))
                    );
                    itemInput.click();
                    itemInput.sendKeys(itemBarcode);
                    Thread.sleep(3000);
                    IBDocPathManager.captureScreenshot(driver,"Enter item"); //Screenshot
                    IBDocPathManager.saveSharedDocument();                               //Screenshot Doc
                    System.out.println("Enter Item screenshot done");
                    itemInput.sendKeys(Keys.ENTER);

                    // Step 12: Enter Quantity Unit
                    WebElement qtyInput = wait.until(
                            ExpectedConditions.presenceOfElementLocated(
                                    By.xpath("//input[@data-component-id='acceptquantity_naturalquantityfield_unit']"))
                    );
                    qtyInput.click();
                    qtyInput.sendKeys(enterQuantityUnit);
                    qtyInput.sendKeys(Keys.ENTER);
                    Thread.sleep(3000);
                    IBDocPathManager.captureScreenshot(driver,"Enter QTY"); //Screenshot
                    IBDocPathManager.saveSharedDocument();                               //Screenshot Doc
                    System.out.println("Enter QTY screenshot done");
                    Thread.sleep(5000);

                    WebElement okButton = wait.until(ExpectedConditions.elementToBeClickable(
                            By.xpath("//button[.//span[text()='Ok']]")
                    ));
                    okButton.click();
                    driver.switchTo().window(secondTabHandle);
                    Thread.sleep(2000);

                    // Step 13: Enter Direct To Pallet
                    WebElement palletInput = wait.until(
                            ExpectedConditions.presenceOfElementLocated(
                                    By.xpath("//input[@placeholder='Scan Pallet']"))
                    );
                    palletInput.click();
                    palletInput.sendKeys(directToPallet);
                    Thread.sleep(3000);
                    IBDocPathManager.captureScreenshot(driver,"Enter Pallet"); //Screenshot
                    IBDocPathManager.saveSharedDocument();                               //Screenshot Doc
                    System.out.println("Enter Pallet screenshot done");
                    palletInput.sendKeys(Keys.ENTER);
                }

                // At the end of this testcase group, release dock door
                WebElement releaseDock = wait.until(
                        ExpectedConditions.presenceOfElementLocated(
                                By.xpath("//button[@data-component-id='action_releasedoor_button']"))
                );
                releaseDock.click();

                // === Step 15: Exit to main menu ===
                WebElement exit_menu = wait.until(
                        ExpectedConditions.elementToBeClickable(
                                By.cssSelector("ion-button[data-component-id='action_exit_button']"))
                );
                exit_menu.click();

                // === Step 16: Confirm popup ===
                WebElement ConfirmButton = wait.until(ExpectedConditions.elementToBeClickable(
                        By.xpath("//button[.//span[text()='Confirm']]")
                ));
                ConfirmButton.click();
                driver.switchTo().window(secondTabHandle);

                // allow small pause between testcases
                Thread.sleep(1000);
            }

        /*
         This is called in Main function

            //Step 17: System directed Pallet Putaway
            Thread.sleep(5000);
            pallet_putaway step4 = new pallet_putaway(driver);
            step4.execute();
        */



        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // helper: return string cell value safely
    private static String getCellValue(Cell cell) {
        if (cell == null) return "";
        DataFormatter formatter = new DataFormatter();
        return formatter.formatCellValue(cell).trim();
    }

    // group rows by Testcase value in sheet (preserves order)
    private Map<String, List<Row>> groupRows(Sheet sheet) {
        Map<String, List<Row>> map = new LinkedHashMap<>();
        if (sheet == null) return map;
        Iterator<Row> it = sheet.iterator();
        if (it.hasNext()) it.next();
        while (it.hasNext()) {
            Row row = it.next();
            if (row == null) continue;
            // try to find Testcase cell by header 'Testcase' or fallback to col 0
            String tc = "";
            Row header = sheet.getRow(0);
            int tcIndex = -1;
            if (header != null) {
                for (int c = 0; c < header.getLastCellNum(); c++) {
                    String h = getCellValue(header.getCell(c));
                    if ("Testcase".equalsIgnoreCase(h)) {
                        tcIndex = c;
                        break;
                    }
                }
            }
            if (tcIndex != -1) {
                Cell c = row.getCell(tcIndex);
                if (c != null) tc = getCellValue(c);
            } else {
                Cell c = row.getCell(0);
                if (c != null) tc = getCellValue(c);
            }
            if (tc == null) tc = "";
            tc = tc.trim();
            if (!tc.matches("TST_\\d+")) continue;
            map.computeIfAbsent(tc, x -> new ArrayList<>()).add(row);
        }
        return map;
    }
}
