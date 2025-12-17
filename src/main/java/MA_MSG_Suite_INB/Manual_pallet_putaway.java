package MA_MSG_Suite_INB;

import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
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

public class Manual_pallet_putaway {

    private WebDriver driver;

    // ✅ Excel data path now comes from central ExcelReaderIB
    private static final String DATA_EXCEL_PATH = ExcelReaderIB.DATA_EXCEL_PATH;

    public Manual_pallet_putaway(WebDriver driver) {
        this.driver = driver;
    }

    // Helper method to safely read Excel cells as String
    private String getCellValueAsString(Cell cell) {
        if (cell == null) return "";
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue().trim();
            case NUMERIC -> String.valueOf((long) cell.getNumericCellValue()).trim();
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue()).trim();
            case FORMULA -> cell.getCellFormula().trim();
            case BLANK, _NONE, ERROR -> "";
        };
    }

    public void execute() throws InterruptedException {
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(20));

        // Step 4: Read Excel
        try (FileInputStream fis = new FileInputStream(DATA_EXCEL_PATH);
             Workbook workbook = WorkbookFactory.create(fis)) {

            Sheet sheet = workbook.getSheet("pallet_putaway");
            if (sheet == null) {
                throw new RuntimeException("Sheet 'pallet_putaway' not found!");
            }

            // If Testcase grouping exists, process per-testcase; otherwise fallback to old behavior
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
            if (headerRow == null) throw new RuntimeException("Sheet 'pallet_putaway' has no header row!");

            int trxnCol = -1, scanContainerCol = -1;
            for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                String colName = getCellValue(headerRow.getCell(i));
                switch (colName) {
                    case "Trxn_Name" -> trxnCol = i;
                    case "ScanContainer" -> scanContainerCol = i;
                }
            }

            if (trxnCol == -1 || scanContainerCol == -1) {
                throw new RuntimeException("Required columns not found in Excel!");
            }

            // Process each testcase group in sheet order
            for (Map.Entry<String, List<Row>> entry : grouped.entrySet()) {
                String testcase = entry.getKey();
                List<Row> rows = entry.getValue();
                if (rows.isEmpty()) continue;

                System.out.println("▶ Processing Testcase: " + testcase + " (" + rows.size() + " rows)");

                // ✅ Step 6–7: Get fixed transaction name (from first row of the group only)
                Row firstRow = rows.get(0);
                String trxnName = getCellValueAsString(firstRow.getCell(trxnCol));

                // Step 6: Enter Trxn_Name in Search
                WebDriverWait searchWait = new WebDriverWait(driver, Duration.ofSeconds(20));
                WebElement newTabSearch = searchWait.until(
                        ExpectedConditions.presenceOfElementLocated(
                                By.xpath("//input[@placeholder='Search']"))
                );
                newTabSearch.click();
                newTabSearch.sendKeys(Keys.chord(Keys.CONTROL, "a"));
                newTabSearch.sendKeys(Keys.BACK_SPACE);
                newTabSearch.sendKeys(trxnName);
                newTabSearch.sendKeys(Keys.ENTER);

                // Step 7: Match result by text and click it
                List<WebElement> allElements = driver.findElements(By.xpath("//*"));
                for (WebElement el : allElements) {
                    try {
                        if (el.isDisplayed() && el.getText().trim().equals(trxnName)) {
                            el.click();
                            break;
                        }
                    } catch (Exception ignored) {}
                }

                // ✅ Step 8–14: Loop through group's pallets (ScanContainer column only)
                int localIndex = 0;
                for (Row currentRow : rows) {
                    localIndex++;
                    if (currentRow == null) continue;

                    String scanContainer = getCellValueAsString(currentRow.getCell(scanContainerCol));
                    if (scanContainer.isEmpty()) {
                        System.out.println("Skipping row " + localIndex + " (missing pallet)");
                        continue;
                    }

                    // Step 8: Enter Pallet
                    WebElement palletInput = wait.until(
                            ExpectedConditions.presenceOfElementLocated(
                                    By.xpath("//input[@placeholder='Scan Container']"))
                    );
                    palletInput.click();
                    palletInput.sendKeys(scanContainer);
                    palletInput.sendKeys(Keys.ENTER);


                    // Add step 8.1 exception if its taking too long then press oK and then proceed further to step 9
                    //<div id="alert-15-msg" class="alert-message sc-ion-alert-md">The server is taking too long to respond. Please try again.</div>
                    //<button type="button" class="alert-button ion-focusable ion-activatable sc-ion-alert-md" tabindex="0">
                    //    <span class="alert-button-inner sc-ion-alert-md">Ok</span><ion-ripple-effect class="sc-ion-alert-md md hydrated"
                    //      role="presentation"><template shadowrootmode="open"></template></ion-ripple-effect>
                    // </button>
                    // Step 9: Capture Drop Zone dynamically
//                    WebElement dropZoneElement = wait.until(
//                            ExpectedConditions.presenceOfElementLocated(
//                                    By.cssSelector("ion-col[data-component-id='acceptdroplocation_barcodetextfield_dropzone']"))
//                    );
//                    String dropZone = dropZoneElement.getText().trim();
//                    System.out.println("Captured Drop Zone (Row " + localIndex + "): " + dropZone);
//
//                    // Step 10: Switch to first tab
//                    String secondTabHandle = driver.getWindowHandle();
//                    String firstTabHandle = "";
//                    Set<String> windows = driver.getWindowHandles();
//                    for (String window : windows) {
//                        if (!window.equals(secondTabHandle)) {
//                            firstTabHandle = window;
//                            break;
//                        }
//                    }
//                    driver.switchTo().window(firstTabHandle);
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
//                    searchBox1.sendKeys("Pick Drop Locations");
//
//                    WebElement pickDropLocations = wait.until(
//                            ExpectedConditions.presenceOfElementLocated(
//                                    By.xpath("//button[@data-component-id='PickDropLocations']"))
//                    );
//                    pickDropLocations.click();
//
//                    // Step 10.1: Enter Drop Zone in filter
//                    WebElement taskMovementZoneFilter = wait.until(
//                            ExpectedConditions.presenceOfElementLocated(
//                                    By.xpath("//ion-input[@data-component-id='TaskMovementZoneId-lookup-dialog-filter-input']//input"))
//                    );
//                    taskMovementZoneFilter.click();
//                    taskMovementZoneFilter.sendKeys(Keys.chord(Keys.CONTROL, "a"));
//                    taskMovementZoneFilter.sendKeys(Keys.BACK_SPACE);
//                    taskMovementZoneFilter.sendKeys(dropZone);
//                    taskMovementZoneFilter.sendKeys(Keys.ENTER);
//
//                    // Step 10.2: Capture Location Barcode dynamically
//                    Thread.sleep(5000);
//                    WebElement locationBarcodeElement = wait.until(
//                            ExpectedConditions.presenceOfElementLocated(
//                                    By.cssSelector("span[data-component-id='LocationBarcode']"))
//                    );
//                    String locationBarcode = locationBarcodeElement.getText().trim();
//                    System.out.println("Captured Location Barcode (Row " + localIndex + "): " + locationBarcode);
//
//                    // Step 11: Switch back to second tab
//                    driver.switchTo().window(secondTabHandle);
//
//                    WebElement dropLocationInput = wait.until(
//                            ExpectedConditions.presenceOfElementLocated(
//                                    By.xpath("//input[@placeholder='Scan Location']"))
//                    );
//                    dropLocationInput.click();
//                    dropLocationInput.sendKeys(locationBarcode);
//                    dropLocationInput.sendKeys(Keys.ENTER);

                    // Step 9: Capture Drop Zone dynamically
                    WebElement dropZoneElement = wait.until(
                            ExpectedConditions.presenceOfElementLocated(
                                    By.cssSelector("ion-col[data-component-id='acceptdroplocation_barcodetextfield_dropzone']"))
                    );
                    String dropZone = dropZoneElement.getText().trim();
                    System.out.println("Captured Drop Zone (Row " + localIndex + "): " + dropZone);

                    // ⭐ NEW — Single API call replaces entire old block
                    String locationBarcode = LocationBarcodeService.getLocationBarcodeByTaskMovementZone(dropZone);

                    if (locationBarcode == null || locationBarcode.isEmpty()) {
                        System.out.println("❌ No LocationBarcode returned for DropZone: " + dropZone);
                    } else {
                        System.out.println("✅ Location Barcode from API: " + locationBarcode);
                    }

                    // Step 11: Use the barcode in Scan Location input
                    WebElement dropLocationInput = wait.until(
                            ExpectedConditions.presenceOfElementLocated(
                                    By.xpath("//input[@placeholder='Scan Location']"))
                    );
                    dropLocationInput.click();
                    dropLocationInput.sendKeys(locationBarcode);
                    dropLocationInput.sendKeys(Keys.ENTER);

                    // Step 12: Enter Pallet for final destination
                    WebElement palletInput1 = wait.until(
                            ExpectedConditions.presenceOfElementLocated(
                                    By.xpath("//input[@placeholder='Scan Container']"))
                    );
                    palletInput1.click();
                    palletInput1.sendKeys(scanContainer);
                    palletInput1.sendKeys(Keys.ENTER);

                    // Step 13: Capture Final Location dynamically
                    WebElement finalLocationElement = wait.until(
                            ExpectedConditions.presenceOfElementLocated(
                                    By.cssSelector("ion-col[data-component-id='acceptlocationforsystemdirectedputaway_barcodetextfield_location']"))
                    );
                    String finalLocation = finalLocationElement.getText().trim();
                    finalLocation = finalLocation.replaceAll("[^a-zA-Z0-9]", "");
                    System.out.println("Captured Final Location (Row " + localIndex + "): " + finalLocation);

                    // Step 14: Send the Final Location for putaway
                    WebElement location = wait.until(
                            ExpectedConditions.presenceOfElementLocated(
                                    By.xpath("//input[@placeholder='Scan Location']"))
                    );
                    location.click();
                    location.sendKeys(finalLocation);
                    location.sendKeys(Keys.ENTER);

                    // Small wait before moving to next pallet
                    Thread.sleep(2000);
                }

                // small pause between testcases
                Thread.sleep(1000);
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
