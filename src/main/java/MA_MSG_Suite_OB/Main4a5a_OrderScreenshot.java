

package MA_MSG_Suite_OB;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.*;

import java.io.*;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.*;
import java.util.NoSuchElementException;

public class Main4a5a_OrderScreenshot {

    // ====== CONFIG ======
    public static WebDriver driver;
    public static  int time=60;
    public static XWPFDocument document = new XWPFDocument();
     // public static String filePath ;
              //= "C:\\Users\\2389120\\IdeaProjects\\msg-runner\\OOdata.xlsx";
      //public static String TESTCASE_VALUE ;
                      //= "TST_001"; // change if needed
    private static final Duration DEFAULT_WAIT = Duration.ofSeconds(5);
    private static final Duration FRAME_SEARCH_TIMEOUT = Duration.ofSeconds(5);

    //
    // ====== MAIN ======
       public static void main(String filePath,String TESTCASE_VALUE,String Sheetname,String env) {
  //  public static void main(String TESTCASE_VALUE, String filePath) throws IOException {
      //  closeExcelIfOpen();

//    String[] args = new String[0];
//    if (args != null && args.length > 0 && args[0] != null && !args[0].trim().isEmpty()) {
//            TESTCASE_VALUE = args[0].trim();
//        }

        List<String> orders;
        try {
            orders = fetchPassedOriginalOrdersForTestcase(TESTCASE_VALUE,filePath,Sheetname);
        } catch (Exception e) {
            System.err.println("Excel read failed: " + e.getMessage());
            e.printStackTrace();
            return;
        }

        if (orders.isEmpty()) {
            System.out.println("No 'Passed' rows found for Testcase: " + TESTCASE_VALUE + ". Exiting.");
            return;
        }
        System.out.println("Found " + orders.size() + " OriginalOrderId(s): " + orders);

        try {
            // WebDriver & Login
            WebDriverManager.chromedriver().setup();
            ChromeOptions options = new ChromeOptions();
            options.addArguments("--start-maximized");
            driver = new ChromeDriver(options);
            driver.manage().window().maximize();
            Main1_URL_Login1 login1 = new Main1_URL_Login1(driver, env);
            login1.execute();
            System.out.println("login done:\n");


            //navigateToOriginalOrdersUI();
            SearchMenu("Original Orders 2.0","originalOrdersVer2");

            // Switch to the frame that contains the Actual app content (ion-app/filters)
            // We look for the OriginalOrderId input host OR ion-app as a reliable marker.
            By appMarker = By.cssSelector("ion-app, ma-original-orders-v2, ion-input[data-component-id='OriginalOrderId']");
            boolean inFrame = switchToFrameContaining(appMarker, FRAME_SEARCH_TIMEOUT);
            System.out.println("Switched to content frame? " + inFrame);

            waitForOverlaysToDisappear();

            // Try to expand filters (only if needed)
            ensureFiltersAndInputVisible();
          //  String docPathLocal = buildDocPath(filePath, TESTCASE_VALUE);
            String docPathLocal = DocPathManager.getOrCreateDocPath(filePath, TESTCASE_VALUE);
            // Process each order
            for (int i = 0; i < orders.size(); i++) {
                String orderId = orders.get(i);
                System.out.println("\n=== Processing (" + (i + 1) + "/" + orders.size() + "): " + orderId + " ===");
                ensureFiltersAndInputVisible();
                boolean entered = enterOriginalOrderId(orderId);
                if (!entered) {
                    System.out.println("Failed to enter order ID. Skipping: " + orderId);
                    continue;
                }

                waitForOverlaysToDisappear();
                applySortIfNeeded();

                System.out.println("Output doc: " + docPathLocal);
                captureScreenshot("Orders");
              //  captureAllCardsScreenshots(document);
               // saveDocument(docPathLocal, document);
                DocPathManager.saveSharedDocument(); // optional: save after each class


                relatedlinks(docPathLocal);
                Thread.sleep(10000);


            }






            System.out.println("\nAll orders processed.");
        } catch (Exception e) {
            System.err.println("Run failed: " + e.getMessage());
            e.printStackTrace();
        } finally {
            if (driver != null) driver.quit();
            try { document.close(); } catch (IOException ignored) {}
        }
    }

    // ====== EXCEL ======
    public static List<String> fetchPassedOriginalOrdersForTestcase(String testcaseValue,String filePath,String Sheetname) {
        final String sheetName = Sheetname;
        List<String> originalOrderIds = new ArrayList<>();
        DataFormatter formatter = new DataFormatter();

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheet(sheetName);
            if (sheet == null) {
                sheet = workbook.getNumberOfSheets() > 0 ? workbook.getSheetAt(0) : null;
            }
            if (sheet == null) throw new IllegalStateException("Sheet not found: " + sheetName);

            int headerRowIdx = findHeaderRowIndexSimple(sheet);
            if (headerRowIdx < 0) throw new IllegalStateException("Header row not found.");
            Row headerRow = sheet.getRow(headerRowIdx);

            Map<String, Integer> headerIndex = buildHeaderIndexExact(headerRow, formatter);
            System.out.println("Detected headers (exact): " + headerIndex);

            Integer tcCol = headerIndex.get("Testcase");
            Integer resultCol = headerIndex.get("Result");
            Integer originalOrderIdCol = headerIndex.get("OriginalOrderId");

            if (tcCol == null) throw new IllegalStateException("Column 'Testcase' not found.");
            if (resultCol == null) throw new IllegalStateException("Column 'Result' not found.");
            if (originalOrderIdCol == null) throw new IllegalStateException("Column 'OriginalOrderId' not found.");

            int firstDataRow = headerRowIdx + 1;
            for (int r = firstDataRow; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;

                String tc = cleanupInvisible(safeFormat(formatter, row.getCell(tcCol)));
                if (!testcaseValue.equals(tc)) continue;

                String result = cleanupInvisible(safeFormat(formatter, row.getCell(resultCol)));
                if (!("Passed".equalsIgnoreCase(result) || "Passed (HTTP 200)".equalsIgnoreCase(result))) {
                    continue;
                }


                String originalOrderId = cleanupInvisible(safeFormat(formatter, row.getCell(originalOrderIdCol)));
                if (!originalOrderId.isEmpty()) originalOrderIds.add(originalOrderId);
            }

            if (originalOrderIds.isEmpty()) {
                System.out.println("No 'Passed' rows found for Testcase: " + testcaseValue);
            } else {
                System.out.println("OriginalOrderId values (Result=Passed) for Testcase " + testcaseValue + ":");
                for (int i = 0; i < originalOrderIds.size(); i++) {
                    System.out.println((i + 1) + ". " + originalOrderIds.get(i));
                }
            }
        } catch (Exception e) {
            throw new RuntimeException("Failed to read Excel and fetch OriginalOrderId values: " + e.getMessage(), e);
        }

        return originalOrderIds;
    }

    private static int findHeaderRowIndexSimple(Sheet sheet) {
        int last = Math.min(sheet.getLastRowNum(), sheet.getFirstRowNum() + 20);
        DataFormatter formatter = new DataFormatter();
        for (int r = sheet.getFirstRowNum(); r <= last; r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            short firstCell = row.getFirstCellNum();
            short lastCell = row.getLastCellNum();
            if (firstCell < 0 || lastCell < 0) continue;
            int nonEmpty = 0;
            for (int c = firstCell; c < lastCell; c++) {
                String value = cleanupInvisible(safeFormat(formatter, row.getCell(c)));
                if (!value.isEmpty()) nonEmpty++;
            }
            if (nonEmpty >= 3) return r;
        }
        return -1;
    }

    private static Map<String, Integer> buildHeaderIndexExact(Row headerRow, DataFormatter formatter) {
        Map<String, Integer> headerIndex = new HashMap<>();
        short firstHeaderCell = headerRow.getFirstCellNum();
        short lastHeaderCell = headerRow.getLastCellNum();
        for (int c = firstHeaderCell; c < lastHeaderCell; c++) {
            String header = cleanupInvisible(safeFormat(formatter, headerRow.getCell(c)));
            if (!header.isEmpty()) headerIndex.putIfAbsent(header, c);
        }
        return headerIndex;
    }

    private static String cleanupInvisible(String s) {
        if (s == null) return "";
        s = s.replace("\uFEFF", "").replace("\u00A0", " ");
        s = s.trim();
        if (s.length() >= 2 && ((s.startsWith("\"") && s.endsWith("\"")) || (s.startsWith("'") && s.endsWith("'")))) {
            s = s.substring(1, s.length() - 1).trim();
        }
        return s;
    }

    private static String safeFormat(DataFormatter formatter, Cell cell) {
        if (cell == null) return "";
        return formatter.formatCellValue(cell).trim();
    }

    // ====== BROWSER / UI ======


    public static void navigateToOriginalOrdersUI() throws InterruptedException {
        WebDriverWait wait = new WebDriverWait(driver, DEFAULT_WAIT);
        JavascriptExecutor js = (JavascriptExecutor) driver;

        sleep(3000);
        waitForOverlaysToDisappear();

        // Open side menu (ion-button with shadow)
        WebElement menuHost = driver.findElement(By.cssSelector("ion-button.menu-toggle-button"));
        // Use JS to click the host (robust even if shadow changes)
        js.executeScript("arguments[0].click();", menuHost);

        // Search menu item
        WebElement searchInput = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@placeholder='Search Menu...']")));
        searchInput.clear();
        searchInput.sendKeys("Original Orders 2.0");

        WebElement originalOrdersBtn = wait.until(ExpectedConditions.elementToBeClickable(By.id("originalOrdersVer2")));
        js.executeScript("arguments[0].click();", originalOrdersBtn);

        sleep(3000);


        // Close menu if present
        try {
           // WebElement closeButton = wait.until(ExpectedConditions.elementToBeClickable(By.id("close-menu-button")));

            WebElement closeButton = driver.findElement(By.id("close-menu-button"));



            closeButton.click();
        } catch (TimeoutException e) {
            System.out.println("Close menu button not found or already closed.");
        }

        waitForOverlaysToDisappear();
    }

    /**
     * Try to ensure filters are open and the OriginalOrderId input is visible.
     * We attempt multiple strategies and bail out early if the input is already present.
     */
    private static void ensureFiltersAndInputVisible() {
        WebDriverWait wait = new WebDriverWait(driver, DEFAULT_WAIT);
        JavascriptExecutor js = (JavascriptExecutor) driver;

        // If the input is already available, skip expansion.
        if (isIonInputPresent("OriginalOrderId", Duration.ofSeconds(10))) {
            System.out.println("OriginalOrderId input found; no need to expand filters.");
            return;
        }

        // Try clickable toggle buttons first (generic strategy)
        try {
            WebElement filterToggle = wait.until(ExpectedConditions.presenceOfElementLocated(
                    By.xpath("(//ion-button[contains(@class,'toggle-button')])[3]")));
            js.executeScript(
                    "const host = arguments[0];" +
                            "const sr = host.shadowRoot;" +
                            "const inner = sr ? sr.querySelector('.button-inner') : host.querySelector('.button-inner');" +
                            "if(inner) inner.click();", filterToggle);
            sleep(500);
        } catch (Exception e) {
            System.out.println("Filter toggle not used: " + e.getMessage());
        }

        // Try a more generic chevron-down (component-ids can vary)
        try {
            WebElement chevronBtn = new WebDriverWait(driver, Duration.ofSeconds(time))
                    .until(ExpectedConditions.presenceOfElementLocated(
                            By.xpath("//ion-button[contains(@data-component-id,'chevron') or contains(@aria-label,'Expand') or contains(., 'Expand')]")));
            js.executeScript("arguments[0].scrollIntoView({block:'center'});", chevronBtn);
            sleep(250);
            js.executeScript("arguments[0].click();", chevronBtn);
            System.out.println("Chevron/Expand button clicked via generic selector.");
        } catch (TimeoutException te) {
            System.out.println("Chevron/Expand button not found—continuing.");
        }

        // Final check: input present?
        if (!isIonInputPresent("OriginalOrderId", Duration.ofSeconds(time))) {
            System.out.println("OriginalOrderId input still not visible after expansion attempts.");
        }
    }

    /**
     * Enters the given Original Order ID.
     * Tries native (inner input) then JS across shadow DOM; also tries locating by label text if component-id differs.
     */
    public static boolean enterOriginalOrderId(String valueToEnter) {
        // 1) Native inner input (if projection to light DOM is available)
        try {
            WebElement innerInputEl = new WebDriverWait(driver, DEFAULT_WAIT)
                    .until(ExpectedConditions.elementToBeClickable(
                            By.xpath("//ion-input[@data-component-id='OriginalOrderId']//input[contains(@class,'native-input')]")));
            ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView({block:'center'});", innerInputEl);
            sleep(150);

            innerInputEl.click();
            innerInputEl.sendKeys(Keys.chord(Keys.CONTROL, "a"));
            innerInputEl.sendKeys(Keys.DELETE);
            innerInputEl.sendKeys(valueToEnter);
            innerInputEl.sendKeys(Keys.ENTER);

            System.out.println("Order ID entered (native sendKeys): " + valueToEnter);
            return true;
        } catch (TimeoutException | ElementClickInterceptedException e) {
            System.out.println("Native interaction failed; trying JS across shadow DOM.");
        }

        // 2) JS: set on ion-input by data-component-id
        if (setIonInputValueByComponentId("OriginalOrderId", valueToEnter)) {
            System.out.println("Order ID set via JS on ion-input[data-component-id='OriginalOrderId']: " + valueToEnter);
            return true;
        }

        // 3) JS: try by label text (e.g., "Original Order Id") in case component-id differs
        if (setIonInputValueByLabel("Original Order Id", valueToEnter)) {
            System.out.println("Order ID set via JS by label 'Original Order Id': " + valueToEnter);
            return true;
        }

        return false;
    }

    /** Checks if an ion-input with given data-component-id is present (any frame context assumed already active). */
    private static boolean isIonInputPresent(String dataComponentId, Duration timeout) {
        try {
            new WebDriverWait(driver, timeout)
                    .until(ExpectedConditions.presenceOfElementLocated(
                            By.cssSelector("ion-input[data-component-id='" + dataComponentId + "']")));
            return true;
        } catch (TimeoutException te) {
            return false;
        }
    }

    /**
     * Sets an ion-input value by its data-component-id using JS (handles shadow DOM).
     */
    private static boolean setIonInputValueByComponentId(String dataComponentId, String value) {
        JavascriptExecutor js = (JavascriptExecutor) driver;
        try {
            Object ok = js.executeScript(
                    "const id = arguments[0], val = arguments[1];" +
                            "const hosts = Array.from(document.querySelectorAll('ion-input')); " +
                            "let target = hosts.find(h => (h.getAttribute('data-component-id')||'').trim() === id); " +
                            "if (!target) return false; " +
                            "const sr = target.shadowRoot; " +
                            "const inputEl = sr ? sr.querySelector('input') : target.querySelector('input'); " +
                            "if (!inputEl) return false; " +
                            "inputEl.value = val; " +
                            "inputEl.dispatchEvent(new Event('input', { bubbles: true })); " +
                            "inputEl.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true })); " +
                            "return true;",
                    dataComponentId, value
            );
            return Boolean.TRUE.equals(ok);
        } catch (JavascriptException je) {
            System.out.println("JS error while setting ion-input by component-id: " + je.getMessage());
            return false;
        }
    }

    /**
     * Sets an ion-input value by nearby label text (fallback if component-id not stable).
     * It scans ion-item elements and looks for text match, then finds nested ion-input.
     */
    private static boolean setIonInputValueByLabel(String labelText, String value) {
        JavascriptExecutor js = (JavascriptExecutor) driver;
        try {
            Object ok = js.executeScript(
                    "const label = arguments[0], val = arguments[1];" +
                            "function text(el){ return (el.textContent||'').trim(); }" +
                            "const items = Array.from(document.querySelectorAll('ion-item')); " +
                            "for (const item of items) {" +
                            "  if (text(item).toLowerCase().includes(label.toLowerCase())) {" +
                            "    const inputHost = item.querySelector('ion-input'); " +
                            "    if (!inputHost) continue; " +
                            "    const sr = inputHost.shadowRoot; " +
                            "    const inputEl = sr ? sr.querySelector('input') : inputHost.querySelector('input'); " +
                            "    if (!inputEl) continue; " +
                            "    inputEl.value = val; " +
                            "    inputEl.dispatchEvent(new Event('input', { bubbles: true })); " +
                            "    inputEl.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true })); " +
                            "    return true; " +
                            "  }" +
                            "}" +
                            "return false;",
                    labelText, value
            );
            return Boolean.TRUE.equals(ok);
        } catch (JavascriptException je) {
            System.out.println("JS error while setting ion-input by label: " + je.getMessage());
            return false;
        }
    }

    /** Optional sort (customize to your UI). Tries a generic sort button or grid header. */
    private static void applySortIfNeeded() {
        WebDriverWait wait = new WebDriverWait(driver, DEFAULT_WAIT);
        JavascriptExecutor js = (JavascriptExecutor) driver;

        // Try a known sort button first
        try {
            WebElement sortBtn = wait.withTimeout(Duration.ofSeconds(5))
                    .until(ExpectedConditions.presenceOfElementLocated(
                            By.cssSelector("ion-button[data-component-id='OriginalOrderVer2-SortButton']")));
            js.executeScript("arguments[0].scrollIntoView({block:'center'});", sortBtn);
            sleep(200);
            js.executeScript("arguments[0].click();", sortBtn);
            waitForOverlaysToDisappear();
            System.out.println("Sort button clicked.");
            return;
        } catch (TimeoutException ignored) {}

        // Otherwise click first grid header if present
        try {
            WebElement firstHeader = wait.withTimeout(Duration.ofSeconds(5))
                    .until(ExpectedConditions.presenceOfElementLocated(
                            By.cssSelector("[role='columnheader'], .ag-header-cell, .grid-header")));
            js.executeScript("arguments[0].scrollIntoView({block:'center'});", firstHeader);
            sleep(150);
            js.executeScript("arguments[0].click();", firstHeader);
            waitForOverlaysToDisappear();
            System.out.println("Generic grid header clicked for sort.");
        } catch (TimeoutException ignored) {
            // No sort element found; skip
        }
    }

    /** Waits for overlays/spinners (Ionic/Manhattan) to disappear. */
    private static void waitForOverlaysToDisappear() {
        WebDriverWait wait = new WebDriverWait(driver, DEFAULT_WAIT);
        try {
            wait.until(ExpectedConditions.invisibilityOfElementLocated(
                    By.cssSelector("manh-overlay-container, ion-loading, .backdrop")));
        } catch (TimeoutException ignored) {}
    }

    /** Recursively search frames and switch into the one that contains the locator. */
    private static boolean switchToFrameContaining(By locator, Duration timeout) {
        long end = System.currentTimeMillis() + timeout.toMillis();
        while (System.currentTimeMillis() < end) {
            try {
                // Try in default content first
                driver.switchTo().defaultContent();
                if (elementExists(locator, Duration.ofSeconds(time))) return true;
                // Search frames recursively
                if (searchFramesRecursive(locator)) return true;
            } catch (Exception ignored) {}

            sleep(500);
        }
        return false;
    }

    private static boolean searchFramesRecursive(By locator) {
        List<WebElement> frames = driver.findElements(By.cssSelector("iframe, frame"));
        for (int i = 0; i < frames.size(); i++) {
            try {
                driver.switchTo().frame(i);
                if (elementExists(locator, Duration.ofSeconds(time))) return true;
                // Recurse deeper
                if (searchFramesRecursive(locator)) return true;
                driver.switchTo().parentFrame();
            } catch (NoSuchFrameException nsf) {
                driver.switchTo().defaultContent();
            } catch (Exception e) {
                driver.switchTo().parentFrame();
            }
        }
        return false;
    }

    private static boolean elementExists(By locator, Duration timeout) {
        try {
            new WebDriverWait(driver, timeout)
                    .until(ExpectedConditions.presenceOfElementLocated(locator));
            return true;
        } catch (TimeoutException te) {
            return false;
        }
    }

    private static void sleep(long millis) {
        try { Thread.sleep(millis); } catch (InterruptedException ie) { Thread.currentThread().interrupt(); }
    }


//    public static void captureScreenshot(String fileName, XWPFDocument document) {
//        try {
//            File srcFile = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
//            try (FileInputStream fis = new FileInputStream(srcFile)) {
//                XWPFParagraph paragraph = document.createParagraph();
//                XWPFRun run = paragraph.createRun();
//                run.setText("Screenshot: " + fileName);
//                run.addBreak();
//                run.addPicture(fis,
//                        Document.PICTURE_TYPE_PNG,
//                        fileName + ".png",
//                        Units.toEMU(500),
//                        Units.toEMU(300)); // taller height to avoid header-only capture
//            }
//            System.out.println("Screenshot added to document: " + fileName);
//        } catch (Exception e) {
//            System.out.println("Error capturing screenshot: " + e.getMessage());
//        }
//    }


    public static void captureScreenshot(String fileName) {
        try {
            File srcFile = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
            try (FileInputStream fis = new FileInputStream(srcFile)) {
                XWPFDocument document = DocPathManager.getSharedDocument();
                XWPFParagraph paragraph = document.createParagraph();
                XWPFRun run = paragraph.createRun();
                run.setText("Screenshot: " + fileName);
                run.addBreak();
                run.addPicture(fis,
                        Document.PICTURE_TYPE_PNG,
                        fileName + ".png",
                        Units.toEMU(500),
                        Units.toEMU(300));
            }
            System.out.println("Screenshot added to document: " + fileName);
        } catch (Exception e) {
            System.out.println("Error capturing screenshot: " + e.getMessage());
        }
    }
    public static void captureAllCardsScreenshots() throws InterruptedException, IOException {
        XWPFDocument document = DocPathManager.getSharedDocument(); // shared doc
        List<WebElement> rows = driver.findElements(By.cssSelector("[role='main'] card-view"));
        int i = 1;
        for (WebElement row : rows) {
            ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView({block:'center'});", row);
            Thread.sleep(500);
            captureScreenshotRow(row, i, document);
            Thread.sleep(800);
            i++;
        }
    }

    public static void captureScreenshotRow(WebElement ele, int i, XWPFDocument document) {
        try {
            File srcFile = ele.getScreenshotAs(OutputType.FILE);
            try (FileInputStream fis = new FileInputStream(srcFile)) {
                XWPFParagraph paragraph = document.createParagraph();
                XWPFRun run = paragraph.createRun();
                run.setText("Card Row Screenshot: " + i);
                run.addBreak();
                run.addPicture(fis, Document.PICTURE_TYPE_PNG, i + ".png", Units.toEMU(500), Units.toEMU(100));
            }
            System.out.println("Row screenshot added: " + i);
        } catch (Exception e) {
            System.out.println("Error capturing row screenshot: " + e.getMessage());
        }
    }




    // Row-level capture to avoid repeating top-of-container visuals
//    public static void captureAllCardsScreenshots(XWPFDocument document) throws InterruptedException, IOException {
//        List<WebElement> rows = driver.findElements(
//                By.cssSelector("[role='main'] card-view"));
//        //("card-view[data-component-id='Card-View'] .card-row.primary[tabindex='0']"));
//        int i = 1;
//        for (WebElement row : rows) {
//            ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView({block:'center'});", row);
//            Thread.sleep(500);
//            captureScreenshotRow(row, i, document);
//            Thread.sleep(800);
//            i++;
//        }
//    }
//
//    public static void captureScreenshotRow(WebElement ele, int i, XWPFDocument document) {
//        try {
//            File srcFile = ele.getScreenshotAs(OutputType.FILE);
//            try (FileInputStream fis = new FileInputStream(srcFile)) {
//                XWPFParagraph paragraph = document.createParagraph();
//                XWPFRun run = paragraph.createRun();
//                run.setText("Card Row Screenshot: " + i);
//                run.addBreak();
//                run.addPicture(fis, Document.PICTURE_TYPE_PNG, i + ".png", Units.toEMU(500), Units.toEMU(100));
//            }
//            System.out.println("Row screenshot added: " + i);
//        } catch (Exception e) {
//            System.out.println("Error capturing row screenshot: " + e.getMessage());
//        }
//    }
//

    public static void saveDocument(String docPath, XWPFDocument document) {
        try (FileOutputStream out = new FileOutputStream(docPath)) {
            document.write(out);
            System.out.println("Document saved at: " + docPath);
        } catch (IOException e) {
            System.out.println("Error saving document: " + e.getMessage());
        }
    }

    public static String buildDocPath(String excelPathStr, String baseName) {
        Path excelPath = Paths.get(excelPathStr);
        Path parent = excelPath.getParent() != null ? excelPath.getParent() : Paths.get(".");
        String stamp = new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date());
        String unique = baseName + "_" + stamp + ".docx";
        return parent.resolve(unique).toString();
    }

    public static void SearchMenu(String Keyword, String id) throws InterruptedException {
        WebDriverWait wait1 = new WebDriverWait(driver, Duration.ofSeconds(time));
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(time));
        JavascriptExecutor js = (JavascriptExecutor) driver;


        try {
            WebElement shadowHost = wait1.until(ExpectedConditions.presenceOfElementLocated(
                    By.cssSelector("ion-button[data-component-id='menu-toggle-button']")
            ));
            SearchContext shadowRoot = (SearchContext) js.executeScript("return arguments[0].shadowRoot", shadowHost);
            WebElement nativeButton = shadowRoot.findElement(By.cssSelector("button.button-native"));
            wait1.until(ExpectedConditions.invisibilityOfElementLocated(By.cssSelector("manh-overlay-container")));
            js.executeScript("arguments[0].click();", nativeButton);
            System.out.println("Menu toggle button clicked.");

        } catch (Exception e) {
            System.err.println("Error: " + e.getMessage());
            e.printStackTrace(System.err);
        }


        try {
            WebElement innerInput = wait.until(ExpectedConditions.presenceOfElementLocated(
                    By.cssSelector("ion-input[data-component-id='search-input'] input.native-input")
            ));
            wait.until(ExpectedConditions.elementToBeClickable(innerInput));
            innerInput.clear();
            innerInput.sendKeys(Keyword);
            System.out.println("✅ Search Done: " + Keyword);
        } catch (Exception e) {
            System.err.println("❌ Error interacting with search input: " + e.getMessage());
            e.printStackTrace();
        }

        try {
            WebElement element = wait.until(
                    ExpectedConditions.elementToBeClickable(By.id(id))
            );
            ((JavascriptExecutor) driver).executeScript(
                    "arguments[0].scrollIntoView({block: 'center'});", element
            );
            ((JavascriptExecutor) driver).executeScript(
                    "arguments[0].click();", element
            );
            System.out.println("Clicked element with id: " + id);
        } catch (Exception e) {
            System.err.println("Failed to click element with id: " + id);
            e.printStackTrace();
        }

        Thread.sleep(5000);


        List<WebElement> closeButtons = driver.findElements(By.id("close-menu-button"));

        if (!closeButtons.isEmpty() && closeButtons.get(0).isDisplayed()) {
            System.out.println("GOT");

            // If present and visible, click it
            closeButtons.get(0).click();
        }




    }


    public static void relatedlinks(String docPathLocal) throws IOException, InterruptedException {

        // Click "Related Links" and then "Order Lines" (robust, with stale retry)
        JavascriptExecutor js = (JavascriptExecutor) driver;
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(time));

        try {
            // Locate the card-view element by its data-component-id
            WebElement cardView = driver.findElement(
                    By.cssSelector("card-view[data-component-id='Card-View']")
            );

// Click on it
            cardView.click();




            // Ensure the card area is scrolled into view (optional, but helps on long pages)
            WebElement cardPanel = wait.until(ExpectedConditions.presenceOfElementLocated(
                    By.cssSelector("card-panel > div > div > card-view > div")));
            js.executeScript("arguments[0].scrollIntoView({block:'center'});", cardPanel);

            // Click the Related Links button
            By relatedLinksBtn = By.cssSelector("button[data-component-id='relatedLinks']");
            WebElement relatedLinks = wait.until(ExpectedConditions.elementToBeClickable(relatedLinksBtn));
            js.executeScript("arguments[0].click();", relatedLinks);
            System.out.println("Related Links button clicked.");

        } catch (StaleElementReferenceException stale) {
            System.out.println("Stale element on Related Links. Retrying...");
            WebElement relatedLinksRetry = new WebDriverWait(driver, Duration.ofSeconds(time))
                    .until(ExpectedConditions.elementToBeClickable(By.cssSelector("button[data-component-id='relatedLinks']")));
            js.executeScript("arguments[0].click();", relatedLinksRetry);
            System.out.println("Related Links button clicked after retry.");
        }

// Click "Order Lines" link
        try {
            WebDriverWait waitShort = new WebDriverWait(driver, Duration.ofSeconds(time));
            // The <a> element contains the text "Order Lines"
            By orderLinesLink = By.xpath("//a[normalize-space()='Order Lines']");
            WebElement orderLines = waitShort.until(ExpectedConditions.elementToBeClickable(orderLinesLink));
            js.executeScript("arguments[0].scrollIntoView({block:'center'});", orderLines);
            js.executeScript("arguments[0].click();", orderLines);
            System.out.println("Order Lines clicked.");
        } catch (StaleElementReferenceException stale2) {
            System.out.println("Stale element on Order Lines. Retrying...");
            WebElement orderLinesRetry = new WebDriverWait(driver, Duration.ofSeconds(time))
                    .until(ExpectedConditions.elementToBeClickable(By.xpath("//a[normalize-space()='Order Lines']")));
            js.executeScript("arguments[0].click();", orderLinesRetry);
            System.out.println("Order Lines clicked after retry.");
        } catch (Exception e) {
            System.out.println("Failed to click Order Lines: " + e.getMessage());
        }


        //Add Screenshot Methods here
        System.out.println("Output doc: " + docPathLocal);

        Thread.sleep(10000);
        captureScreenshot("Orders");
        captureAllCardsScreenshots();

        DocPathManager.saveSharedDocument(); // optional: save after each class


       // saveDocument(docPathLocal, document);
        System.out.println("Output doc: " + docPathLocal+" Done");


// Click "Original Orders 2.0" (handles truncated text like "Origina...ders 2.0" and stale elements)
//        JavascriptExecutor js = (JavascriptExecutor) driver;
//        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));

        By originalOrdersLink = By.cssSelector("a[title='Original Orders 2.0']");
        By originalOrdersLi = By.cssSelector("li[data-component-id*='Origina'][data-component-id*='ders2.0']");

// Try direct <a> by title first
        try {
            WebElement link = wait.until(ExpectedConditions.elementToBeClickable(originalOrdersLink));
            js.executeScript("arguments[0].scrollIntoView({block:'center'});", link);
            js.executeScript("arguments[0].click();", link);
            System.out.println("Original Orders 2.0 clicked via <a>.");
        } catch (StaleElementReferenceException se) {
            System.out.println("Stale element on <a> title 'Original Orders 2.0'. Retrying...");
            WebElement linkRetry = new WebDriverWait(driver, Duration.ofSeconds(time))
                    .until(ExpectedConditions.elementToBeClickable(originalOrdersLink));
            js.executeScript("arguments[0].click();", linkRetry);
            System.out.println("Original Orders 2.0 clicked via <a> after retry.");
        } catch (Exception e1) {
            // Fallback: use the <li> with truncated component-id or visible text contains
            try {
                // If <li> is clickable, click it; else find an <a> under it
                WebElement li = wait.until(ExpectedConditions.presenceOfElementLocated(originalOrdersLi));
                js.executeScript("arguments[0].scrollIntoView({block:'center'});", li);

                // Prefer the anchor inside the li if present
                WebElement anchor;
                try {
                    anchor = li.findElement(By.cssSelector("a[title='Original Orders 2.0'], a"));
                } catch (NoSuchElementException nse) {
                    anchor = wait.until(ExpectedConditions.elementToBeClickable(
                            By.xpath("//a[contains(normalize-space(.), 'Original Orders 2.0') or contains(normalize-space(.), 'Origina') and contains(normalize-space(.), 'ders 2.0')]")
                    ));
                }

                wait.until(ExpectedConditions.elementToBeClickable(anchor));
                js.executeScript("arguments[0].click();", anchor);
                System.out.println("Original Orders 2.0 clicked via fallback.");
            } catch (StaleElementReferenceException se2) {
                System.out.println("Stale element on fallback. Retrying...");
                WebElement anchorRetry = new WebDriverWait(driver, Duration.ofSeconds(time))
                        .until(ExpectedConditions.elementToBeClickable(
                                By.xpath("//a[@title='Original Orders 2.0' or (contains(normalize-space(.),'Origina') and contains(normalize-space(.),'ders 2.0'))]")
                        ));
                js.executeScript("arguments[0].click();", anchorRetry);
                System.out.println("Original Orders 2.0 clicked after fallback retry.");
            } catch (Exception e2) {
                System.out.println("Failed to click Original Orders 2.0: " + e2.getMessage());
            }
        }


    }


}


