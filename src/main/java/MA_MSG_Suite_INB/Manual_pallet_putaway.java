package MA_MSG_Suite_INB;

import MA_MSG_Suite_OB.DocPathManager;
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
    public static String docPathLocal;

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

        try (FileInputStream fis = new FileInputStream(DATA_EXCEL_PATH);
             Workbook workbook = WorkbookFactory.create(fis)) {

            Sheet sheet = workbook.getSheet("pallet_putaway");
            if (sheet == null) {
                throw new RuntimeException("Sheet 'pallet_putaway' not found!");
            }

            Map<String, List<Row>> grouped = groupRows(sheet);
            if (grouped.isEmpty()) {
                List<Row> allRows = new ArrayList<>();
                Iterator<Row> itAll = sheet.iterator();
                if (itAll.hasNext()) itAll.next();
                while (itAll.hasNext()) {
                    Row r = itAll.next();
                    if (r != null) allRows.add(r);
                }
                grouped.put("ALL_ROWS", allRows);
            }

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

            for (Map.Entry<String, List<Row>> entry : grouped.entrySet()) {
                String testcase = entry.getKey();
                List<Row> rows = entry.getValue();
                if (rows.isEmpty()) continue;

                System.out.println("▶ Processing Testcase: " + testcase + " (" + rows.size() + " rows)");
                docPathLocal = IBDocPathManager.getOrCreateDocPath(ExcelReaderIB.DOC_FILEPATH, testcase);
                System.out.println("Path" + docPathLocal);

                Row firstRow = rows.get(0);
                String trxnName = getCellValueAsString(firstRow.getCell(trxnCol));

                WebDriverWait searchWait = new WebDriverWait(driver, Duration.ofSeconds(20));
                WebElement newTabSearch = searchWait.until(
                        ExpectedConditions.elementToBeClickable(
                                By.xpath("//input[@placeholder='Search']"))
                );
                newTabSearch.click();
                newTabSearch.sendKeys(Keys.chord(Keys.CONTROL, "a"));
                newTabSearch.sendKeys(Keys.BACK_SPACE);
                newTabSearch.sendKeys(trxnName);
                Thread.sleep(3000);
                IBDocPathManager.captureScreenshot(driver, "Enter Transaction");
                IBDocPathManager.saveSharedDocument();
                newTabSearch.sendKeys(Keys.ENTER);

                List<WebElement> allElements = driver.findElements(By.xpath("//*"));
                for (WebElement el : allElements) {
                    try {
                        if (el.isDisplayed() && el.getText().trim().equals(trxnName)) {
                            el.click();
                            break;
                        }
                    } catch (Exception ignored) {
                    }
                }

                int localIndex = 0;
                for (Row currentRow : rows) {
                    localIndex++;
                    if (currentRow == null) continue;

                    String scanContainer = getCellValueAsString(currentRow.getCell(scanContainerCol));
                    if (scanContainer.isEmpty()) continue;

                    WebElement palletInput = wait.until(
                            ExpectedConditions.presenceOfElementLocated(
                                    By.xpath("//input[@placeholder='Scan Container']"))
                    );
                    palletInput.click();
                    palletInput.sendKeys(scanContainer);
                    Thread.sleep(3000);
                    IBDocPathManager.captureScreenshot(driver, "Enter Pallet");
                    IBDocPathManager.saveSharedDocument();
                    palletInput.sendKeys(Keys.ENTER);

                    WebElement dropZoneElement = wait.until(
                            ExpectedConditions.presenceOfElementLocated(
                                    By.cssSelector("ion-col[data-component-id='acceptdroplocation_barcodetextfield_dropzone']"))
                    );
                    String dropZone = dropZoneElement.getText().trim();

                    String locationBarcode = LocationBarcodeService.getLocationBarcodeByTaskMovementZone(dropZone);

                    WebElement dropLocationInput = wait.until(
                            ExpectedConditions.presenceOfElementLocated(
                                    By.xpath("//input[@placeholder='Scan Location']"))
                    );
                    dropLocationInput.click();
                    dropLocationInput.sendKeys(locationBarcode);
                    Thread.sleep(3000);
                    IBDocPathManager.captureScreenshot(driver, "Enter Drop location");
                    IBDocPathManager.saveSharedDocument();
                    dropLocationInput.sendKeys(Keys.ENTER);

                    WebElement palletInput1 = wait.until(
                            ExpectedConditions.presenceOfElementLocated(
                                    By.xpath("//input[@placeholder='Scan Container']"))
                    );
                    palletInput1.click();
                    palletInput1.sendKeys(scanContainer);
                    Thread.sleep(3000);
                    IBDocPathManager.captureScreenshot(driver, "Enter Container");
                    IBDocPathManager.saveSharedDocument();
                    palletInput1.sendKeys(Keys.ENTER);

                    WebElement finalLocationElement = wait.until(
                            ExpectedConditions.presenceOfElementLocated(
                                    By.cssSelector("ion-col[data-component-id='acceptlocationforsystemdirectedputaway_barcodetextfield_location']"))
                    );
                    String finalLocation = finalLocationElement.getText().trim().replaceAll("[^a-zA-Z0-9]", "");

                    WebElement location = wait.until(
                            ExpectedConditions.presenceOfElementLocated(
                                    By.xpath("//input[@placeholder='Scan Location']"))
                    );
                    location.click();
                    location.sendKeys(finalLocation);
                    Thread.sleep(3000);
                    IBDocPathManager.captureScreenshot(driver, "Enter finalLocation");
                    IBDocPathManager.saveSharedDocument();
                    location.sendKeys(Keys.ENTER);

                    Thread.sleep(2000);
                }

                // === Step 15: Exit to main menu ===
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

                // small pause between testcases
                Thread.sleep(1000);
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static String getCellValue(Cell cell) {
        if (cell == null) return "";
        DataFormatter formatter = new DataFormatter();
        return formatter.formatCellValue(cell).trim();
    }

    private Map<String, List<Row>> groupRows(Sheet sheet) {
        Map<String, List<Row>> map = new LinkedHashMap<>();
        if (sheet == null) return map;

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
        if (it.hasNext()) it.next();
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
