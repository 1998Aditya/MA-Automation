package MA_MSG_Suite_INB;

import MA_MSG_Suite_OB.DocPathManager;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openqa.selenium.*;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.List;
import java.util.Set;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Manual_LPN_rcv {

    //private WebDriver driver;
    public static WebDriver driver;
    public static String docPathLocal ;


    // ✅ Excel data path now comes from central ExcelReaderIB
    private static final String DATA_EXCEL_PATH = ExcelReaderIB.DATA_EXCEL_PATH;

    public Manual_LPN_rcv(WebDriver driver) {
        this.driver = driver;
    }

    public void execute() throws InterruptedException {
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(50));

        try {
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
                    secondTabHandle = window;
                    break;
                }
            }

            // Step 4: Read Excel
            try (FileInputStream fis = new FileInputStream(DATA_EXCEL_PATH);
                 Workbook workbook = WorkbookFactory.create(fis)) {

                Sheet sheet = workbook.getSheet("LPN_rcv");
                if (sheet == null) {
                    throw new RuntimeException("Sheet 'LPN_rcv' not found!");
                }

                // If Testcase grouping exists, process per testcase; otherwise fallback to old behavior
                Map<String, List<Row>> grouped = groupRows(sheet);
                if (grouped.isEmpty()) {
                    // fallback: process all rows in original order as single group
                    List<Row> allRows = new ArrayList<>();
                    Iterator<Row> itAll = sheet.iterator();
                    if (itAll.hasNext()) itAll.next(); // skip header
                    while (itAll.hasNext()) {
                        Row r = itAll.next();
                        if (r != null) allRows.add(r);
                    }
                    grouped.put("ALL_ROWS", allRows);
                }

                // Get column indexes from header row
                Row headerRow = sheet.getRow(0);
                if (headerRow == null) throw new RuntimeException("Sheet 'LPN_rcv' has no header row!");

                int trxnCol = -1, dockDoorCol = -1, asnCol = -1, lpnCol = -1, directToPalletCol = -1;
                for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                    String colName = getCellValue(headerRow.getCell(i));
                    switch (colName) {
                        case "Trxn_Name" -> trxnCol = i;
                        case "Dock Door" -> dockDoorCol = i;
                        case "ASN" -> asnCol = i;
                        case "LPN" -> lpnCol = i;
                        case "Direct_To_Pallet" -> directToPalletCol = i;
                    }
                }

                if (trxnCol == -1 || dockDoorCol == -1 || asnCol == -1 ||
                        lpnCol == -1 || directToPalletCol == -1) {
                    throw new RuntimeException("Required columns not found in Excel!");
                }

                // Process each testcase group in sheet order
                for (Map.Entry<String, List<Row>> entry : grouped.entrySet()) {
                    String testcase = entry.getKey();
                    List<Row> rows = entry.getValue();
                    if (rows.isEmpty()) continue;

                    System.out.println("▶ Processing Testcase: " + testcase + " (" + rows.size() + " rows)");
                    docPathLocal = IBDocPathManager.getOrCreateDocPath(ExcelReaderIB.DOC_FILEPATH, testcase);//Screenshot
                    System.out.println("Path"+docPathLocal); //Screenshot Doc
                    // Step 5: Enter Trxn_Name - pick from first row in this testcase
                    Row firstRow = rows.get(0);
                    String trxnName = getCellValue(firstRow.getCell(trxnCol));
                    String dockDoor = getCellValue(firstRow.getCell(dockDoorCol));

                    WebElement newTabSearch = wait.until(
                            ExpectedConditions.presenceOfElementLocated(
                                    By.xpath("//input[@placeholder='Search']"))
                    );
                    Thread.sleep(1000);
                    newTabSearch.click();
                    newTabSearch.sendKeys(Keys.chord(Keys.CONTROL, "a"));
                    newTabSearch.sendKeys(Keys.BACK_SPACE);
                    newTabSearch.sendKeys(trxnName);
                    Thread.sleep(3000);
                    IBDocPathManager.captureScreenshot(driver,"ENTER TRANSACTION"); //Screenshot
                    IBDocPathManager.saveSharedDocument();                               //Screenshot Doc
                    System.out.println("ENTER TRANSACTION screenshot done");
                    Thread.sleep(3000);
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
//                    WebElement dockDoorInput = wait.until(
//                            ExpectedConditions.presenceOfElementLocated(
//                                    By.xpath("//input[@placeholder='Scan Dock Door']"))
//                    );
//                    dockDoorInput.click();
//                    dockDoorInput.sendKeys(dockDoor);
//                    dockDoorInput.sendKeys(Keys.ENTER);

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

                    // === Loop: Group rows by ASN ===
                    String previousASN = "";
                    for (Row row : rows) {
                        if (row == null) continue;

                        String asn = getCellValue(row.getCell(asnCol));
                        String lpn = getCellValue(row.getCell(lpnCol));
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

                        // Step 10: Click OK on popup
                        WebElement okButton = wait.until(ExpectedConditions.elementToBeClickable(
                                By.xpath("//button[.//span[text()='Ok']]")
                        ));
                        okButton.click();
                        driver.switchTo().window(secondTabHandle);
                        Thread.sleep(2000);

                        // Step 11: Enter Direct To Pallet
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

                    // Step 12: Release dock door at the very end
                    WebElement releaseDock = wait.until(
                            ExpectedConditions.presenceOfElementLocated(
                                    By.xpath("//button[@data-component-id='action_releasedoor_button']"))
                    );
                    releaseDock.click();

                    // === Step 13: Exit to main menu ===
                    WebElement exit_menu = wait.until(
                            ExpectedConditions.elementToBeClickable(
                                    By.cssSelector("ion-button[data-component-id='action_exit_button']"))
                    );
                    exit_menu.click();

                    // === Step 14: Confirm popup ===
                    WebElement ConfirmButton = wait.until(ExpectedConditions.elementToBeClickable(
                            By.xpath("//button[.//span[text()='Confirm']]")
                    ));
                    ConfirmButton.click();
                    driver.switchTo().window(secondTabHandle);

                    // small pause between testcases
                    Thread.sleep(1000);
                }

                /*
                This is called in Main function
                // 5. Step 4: System directed Pallet Putaway
                Thread.sleep(5000);
                pallet_putaway step4 = new pallet_putaway(driver);
                step4.execute();
                */
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // ✅ Utility for reading cell values
    private static String getCellValue(Cell cell) {
        if (cell == null) return "";
        DataFormatter formatter = new DataFormatter();
        return formatter.formatCellValue(cell).trim();
    }

    // group rows by Testcase column value (preserve order)
    private Map<String, List<Row>> groupRows(Sheet sheet) {
        Map<String, List<Row>> map = new LinkedHashMap<>();
        if (sheet == null) return map;

        // find Testcase column index from header
        Row header = sheet.getRow(0);
        int tcIndex = -1;
        if (header != null) {
            for (int i = 0; i < header.getLastCellNum(); i++) {
                String h = getCellValue(header.getCell(i));
                if ("Testcase".equalsIgnoreCase(h)) {
                    tcIndex = i;
                    break;
                }
            }
        }

        Iterator<Row> it = sheet.iterator();
        if (it.hasNext()) it.next(); // skip header
        while (it.hasNext()) {
            Row row = it.next();
            if (row == null) continue;
            Cell c = (tcIndex != -1) ? row.getCell(tcIndex) : row.getCell(0);
            if (c == null) continue;
            String tc = getCellValue(c).trim();
            if (!tc.matches("TST_\\d+")) continue;
            map.computeIfAbsent(tc, x -> new ArrayList<>()).add(row);
        }
        return map;
    }


}
