//package MA_MSG_Suite_OB;
//
//import io.github.bonigarcia.wdm.WebDriverManager;
//import org.apache.poi.ss.usermodel.*;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//import org.openqa.selenium.*;
//import org.openqa.selenium.chrome.ChromeDriver;
//import org.openqa.selenium.chrome.ChromeOptions;
//import org.openqa.selenium.support.ui.ExpectedConditions;
//import org.openqa.selenium.support.ui.WebDriverWait;
//
//import java.io.FileInputStream;
//import java.time.Duration;
//import java.util.ArrayList;
//
//
//public class Main9_Palletisation {
//    public static String location;// = "CONHUBALICANTE";
//    public static String pallet;// = "CONHUBALICANTE";
//    public static WebDriver driver;
//    public static void main(String filePath, String testcase, String env) throws InterruptedException {
//        WebDriverManager.chromedriver().setup();
//        ChromeOptions options = new ChromeOptions();
//        options.addArguments("--start-maximized");
//
//        driver = new ChromeDriver(options);
//        driver.manage().window().maximize();
//        Main1_URL_Login1 login1 = new Main1_URL_Login1(driver, env);
//        login1.execute();
//        System.out.println("login done:\n");
//        SearchMenuWM("WM Mobile","WMMobile");
//        //SearchInWmMobile("JD OB Putaway To Staging", "jdobputawaytostaging");
//        //Whatever present in Column E it will pack
//        // OutboundPutaway(filePath,testcase);
//
//    }
//
//    public static void SearchMenuWM(String Keyword, String id)  {
//        WebDriverWait wait1 = new WebDriverWait(driver, Duration.ofSeconds(time));
//        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(time));
//        JavascriptExecutor js = (JavascriptExecutor) driver;
//
//
//
//
//        int maxRetries = 6; // Try up to 2 times (1 initial + 1 retry)
//        for (int attempt = 1; attempt <= maxRetries; attempt++) {
//            System.out.println("⏳ Waiting 10 seconds before refreshing...");
//
//            try {
//                Thread.sleep(10000); // Wait 1 minute
//            } catch (InterruptedException e) {
//                throw new RuntimeException(e);
//            }
//
////                // 🔄 Click the refresh button inside shadow DOM
////                WebElement refreshHost = driver.findElement(By.cssSelector("ion-button.refresh-button"));
////                js = (JavascriptExecutor) driver;
////                WebElement refreshButton = (WebElement) js.executeScript(
////                        "return arguments[0].shadowRoot.querySelector('button.button-native')", refreshHost);
////                refreshButton.click();
////                System.out.println("🔄 Refresh button clicked.");
//
//            // Locate using data-component-id
//            WebElement refreshBtn = wait.until(
//                    ExpectedConditions.elementToBeClickable(
//                            By.cssSelector("ion-button[data-component-id='refresh']")
//                    )
//            );
//
//            // Click the button
//            refreshBtn.click();
//
//            // Optional: verify action or add logging
//            System.out.println("Refresh button clicked successfully!");
//
//            // Optional: wait for UI to settle
//            try {
//                Thread.sleep(3000);
//            } catch (InterruptedException e) {
//                throw new RuntimeException(e);
//            }
//
//            try {
//                WebElement shadowHost = wait1.until(ExpectedConditions.presenceOfElementLocated(
//                        By.cssSelector("ion-button[data-component-id='menu-toggle-button']")
//                ));
//                SearchContext shadowRoot = (SearchContext) js.executeScript("return arguments[0].shadowRoot", shadowHost);
//                WebElement nativeButton = shadowRoot.findElement(By.cssSelector("button.button-native"));
//
//// wait for overlay to disappear
//                wait1.until(ExpectedConditions.invisibilityOfElementLocated(By.cssSelector("manh-overlay-container")));
//
//// click via JS to avoid interception
//                js.executeScript("arguments[0].click();", nativeButton);
//
//                System.out.println("Menu toggle button clicked.");
//
//
//
//
//                break;
//            }catch (Exception e) {
//                System.err.println("Error: " + e.getMessage());
//                e.printStackTrace(System.err);
//
//
//            }
//
//
//        }
//
//
//
//
//        try {
//            // Locate the inner input directly under ion-input
//            WebElement innerInput = wait.until(ExpectedConditions.presenceOfElementLocated(
//                    By.cssSelector("ion-input[data-component-id='search-input'] input.native-input")
//            ));
//
//            wait.until(ExpectedConditions.elementToBeClickable(innerInput));
//
//            innerInput.clear();
//            innerInput.sendKeys(Keyword);
//            System.out.println("✅ Search Done: " + Keyword);
//
//        } catch (Exception e) {
//            System.err.println("❌ Error interacting with search input: " + e.getMessage());
//            e.printStackTrace();
//        }
//        try {
//
//
//            // Wait for the button to be present and visible
//            WebElement element = wait.until(ExpectedConditions.presenceOfElementLocated(
//                    By.cssSelector("button#wmMobile[data-component-id=" + id + "]")
//            ));
//
//
//            ((JavascriptExecutor) driver).executeScript(
//                    "arguments[0].scrollIntoView({block: 'center'});", element
//            );
//
//            ((JavascriptExecutor) driver).executeScript(
//                    "arguments[0].click();", element
//            );
//
//            System.out.println("Clicked element with id: " + id);
//            ArrayList<String> tabs = new ArrayList<>(driver.getWindowHandles());
//            driver.switchTo().window(tabs.get(1));
//
//        } catch (Exception e) {
//
//            System.err.println("Failed to click element with id: " + id);
//            e.printStackTrace();
//
//        }
//
//
//    }
//
//
//    static void SearchInWmMobile(String Transaction, String ComponentId)  {
//        JavascriptExecutor js = (JavascriptExecutor) driver;
//        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(time));
//
//
//        try {
//
//
//            WebElement searchInput = wait.until(ExpectedConditions.elementToBeClickable(
//                    By.cssSelector("ion-searchbar[data-component-id='search'] input[type='search']")));
//
//            // Clear any existing text
//            searchInput.clear();
//
//            // Type the search text
//            searchInput.sendKeys(Transaction);
//
//            // Optionally, press ENTER if the search requires submission
//            searchInput.sendKeys(Keys.ENTER);
//
//
////            WebElement searchInput1 = wait.until(ExpectedConditions.elementToBeClickable(
////                    By.cssSelector("input.searchbar-input[placeholder='Search']")));
//////                    By.xpath("//input[@type='search' and @placeholder='Search']")));
////            searchInput1.click();
////            searchInput1.clear();
////            Thread.sleep(3000);
////            searchInput1.sendKeys(Transaction);
//        } catch (Exception e) {
//            System.err.println("❌ Error in "+Transaction + e.getMessage());
//            e.printStackTrace();
//        }
//        // Locate the ion-label using its data-component-id
//
//        WebElement labelElement = wait.until(ExpectedConditions.elementToBeClickable(
//                By.cssSelector("ion-label[data-component-id='" + ComponentId + "']")
//        ));
//
//        // Scroll into view to ensure it's interactable
//        ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", labelElement);
//
//        // Click using JavaScript (in case native click doesn't work)
//        ((JavascriptExecutor) driver).executeScript("arguments[0].click();", labelElement);
//
//        System.out.println("Clicked on" +Transaction+" label.");
//
//
//    }
//
//    public static void OutboundPutaway(String filePath,String testcase) throws InterruptedException {
//        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(time));
//        JavascriptExecutor js = (JavascriptExecutor) driver;
//
//        try (FileInputStream fis = new FileInputStream(filePath);
//             Workbook workbook = new XSSFWorkbook(fis)) {
//            Sheet sheet = workbook.getSheet("Outbound");
//
//            // Get Row 1 (index 0 because POI is zero-based)
//            Row row = sheet.getRow(0);
//
//            // Map cells to variables
//            String pallet = row.getCell(0).getStringCellValue();  // Row 1, Col 1
//            String StaticRoute = row.getCell(1).getStringCellValue();  // Row 1, Col 2
//            String location = row.getCell(2).getStringCellValue();  // Row 1, Col 3
//            //   int quantity = (int) row.getCell(3).getNumericCellValue(); // Row 1, Col 4
//
//            // Close resources
//            workbook.close();
//            fis.close();
//            System.out.println("Fetching data done ");
//            //return pallet;
//        }
//        catch (Exception e) {
//            System.err.println("❌ Error in Fetching data " + e.getMessage());
//            e.printStackTrace();
//        }
//        //  return pallet;
//
//
//
//
//
//
////        try {
////            Thread.sleep(3000);
////            WebElement searchInput1 = wait.until(ExpectedConditions.elementToBeClickable(
////                    By.cssSelector("input.searchbar-input[placeholder='Search']")));
//////                    By.xpath("//input[@type='search' and @placeholder='Search']")));
////            searchInput1.click();
////            searchInput1.clear();
////            Thread.sleep(3000);
////            searchInput1.sendKeys("JD OB Putaway To Staging");
////        } catch (Exception e) {
////            System.err.println("❌ Error in JD OB Putaway To Staging " + e.getMessage());
////            e.printStackTrace();
////        }
//        // Wait for the page to load (replace with WebDriverWait for production use)
////        Thread.sleep(1000);
////
////        // Locate the ion-label using its data-component-id
////        WebElement labelElement = driver.findElement(
////                By.cssSelector("ion-label[data-component-id='jdobputawaytostaging']")
////        );
////
////        // Scroll into view to ensure it's interactable
////        ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", labelElement);
////
////        // Click using JavaScript (in case native click doesn't work)
////        ((JavascriptExecutor) driver).executeScript("arguments[0].click();", labelElement);
////
////        System.out.println("Clicked on 'JD OB Putaway To Staging' label.");
//
//
//// Optional: wait briefly to ensure stability
//        Thread.sleep(2000); // Or use WebDriverWait if preferred
//
//
//        // Locate the input field using its data-component-id
//        WebElement inputField = driver.findElement(
//                By.cssSelector("input[data-component-id='acceptcontainer_barcodetextfield_scancontainer']")
//        );
//
//        // Clear any existing text and enter the pallet value
//        inputField.clear();
//        inputField.sendKeys(pallet);
//
//        System.out.println("Entered pallet value into Scan Container field: " + pallet);
//
//
//        // Wait until the input field is visible and interactable
//// Try to locate and click the OK button
//        try {
//            inputField.sendKeys(Keys.ENTER);
////            WebElement goButton4 = driver.findElement(
////                    By.cssSelector("ion-button[data-component-id='acceptcontainer_barcodetextfield_go']")
////            );
////
////            // Scroll into view
////            ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", goButton4);
////
////            // Click using JavaScript
////            ((JavascriptExecutor) driver).executeScript("arguments[0].click();", goButton4);
//
//            System.out.println("Clicked on go button.");
//        } catch ( ElementClickInterceptedException e) {
//            System.out.println("Not worked");
//        }
//
//
//
//        Thread.sleep(5000);
//
//        // Try to find scan location input
//        try {
//            WebElement inputField10 = driver.findElement(
//                    By.cssSelector("input[data-component-id='acceptlocation_barcodetextfield_scanlocation']")
//            );
//            inputField10.clear();
//            inputField10.sendKeys(location);
//            System.out.println("Entered scan location: " + location);
//        } catch (Exception e) {
//            // Handle alert and enter destination location instead
//            try {
//                WebElement okButton = driver.findElement(
//                        By.xpath("//div[contains(@class,'alert-button-group')]//button[.//span[text()='Ok']]")
//                );
//                ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", okButton);
//                ((JavascriptExecutor) driver).executeScript("arguments[0].click();", okButton);
//                System.out.println("Clicked on 'Ok' button in alert.");
//            } catch (ElementClickInterceptedException ex) {
//                System.out.println("Pallet is invalid OR  Click manually");
//                Thread.sleep(10000);
//                System.out.println("10 sec remaining");
//                Thread.sleep(10000);
//            }
//            Thread.sleep(4000);
//
//            WebElement inputField11 = driver.findElement(
//                    By.cssSelector("input[data-component-id='acceptlocation_barcodetextfield_destinationlocation']")
//            );
//            inputField11.clear();
//            inputField11.sendKeys(location);
//            System.out.println("Entered location value: " + location);
//            Thread.sleep(5000);
//
//            inputField11.sendKeys(Keys.ENTER);
//
//
//        }
//
//
//
////
//        try {
//            // inputField11.sendKeys(Keys.ENTER);
////            // Locate the ion-button using its data-component-id
////            WebElement goButton5 = driver.findElement(
////                    By.cssSelector("ion-button[data-component-id='acceptlocation_barcodetextfield_go']")
////            );
////
////            // Scroll into view to ensure it's interactable
////            ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", goButton5);
////
////            // Click using JavaScript (in case native click doesn't work)
////            ((JavascriptExecutor) driver).executeScript("arguments[0].click();", goButton5);
////
////            System.out.println("Clicked on 'Go' button for location barcode.");
//
//
//
//            Thread.sleep(3000); // brief pause between clicks
//            // Wait until the ion-buttons container is present
//
//            WebElement backButton = wait.until(ExpectedConditions.presenceOfElementLocated(
//                    By.cssSelector("ion-buttons[data-component-id='action_back_button']")
//            ));
//
//// Scroll into view to ensure visibility
//            ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", backButton);
//
//// Click the button twice using JavaScript
//
//            // js.executeScript("arguments[0].click();", backButton);
//
//            js.executeScript("arguments[0].click();", backButton);
//
//            System.out.println("Back button clicked successfully. 02");
//        } catch (ElementClickInterceptedException e) {
//            System.out.println("Go Not worked");
//            System.out.println(" waiting for Back button clicked ");
//            Thread.sleep(3000); // brief pause between clicks
//            // Wait until the ion-buttons container is present
//
//            WebElement backButton = wait.until(ExpectedConditions.presenceOfElementLocated(
//                    By.cssSelector("ion-buttons[data-component-id='action_back_button']")
//            ));
//
//
//
//        }
//
//        Thread.sleep(3000);
//
//
//        //  public static void tabswitch()
//
//        // Get all open window handles
//        ArrayList<String> tabs = new ArrayList<>(driver.getWindowHandles());
//
//// Switch to the first tab
//        driver.switchTo().window(tabs.get(0));
//
//// Optional: Try to bring the tab to the front using JavaScript
//        // JavascriptExecutor js = (JavascriptExecutor) driver;
//        js.executeScript("window.focus();");
//
//        System.out.println("✅ Switched to the first tab and attempted to bring it to the front.");
//
//
//    }
//
//    /**
//     * Safely get a trimmed String from a cell.
//     * Returns null for null/blank cells.
//     */
//    private static String getCellString(Cell cell) {
//        if (cell == null) return null;
//
//        switch (cell.getCellType()) {
//            case STRING:
//                String s = cell.getStringCellValue();
//                return (s == null) ? null : s.trim();
//            case NUMERIC:
//                // Convert numeric to string without scientific format
//                return new java.text.DecimalFormat("#.################").format(cell.getNumericCellValue()).trim();
//            case BOOLEAN:
//                return String.valueOf(cell.getBooleanCellValue()).trim();
//            case FORMULA:
//                // Evaluate cached value type
//                CellType cached = cell.getCachedFormulaResultType();
//                if (cached == CellType.STRING) {
//                    String fs = cell.getStringCellValue();
//                    return (fs == null) ? null : fs.trim();
//                } else if (cached == CellType.NUMERIC) {
//                    return new java.text.DecimalFormat("#.################").format(cell.getNumericCellValue()).trim();
//                } else if (cached == CellType.BOOLEAN) {
//                    return String.valueOf(cell.getBooleanCellValue()).trim();
//                } else {
//                    return null;
//                }
//            case BLANK:
//            case _NONE:
//            case ERROR:
//            default:
//                return null;
//        }
//    }
//
////// Scroll into view to ensure visibility
////            ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", backButton);
////
////// Click the button twice using JavaScript
////
////            // js.executeScript("arguments[0].click();", backButton);
////            Thread.sleep(300); // brief pause between clicks
////            js.executeScript("arguments[0].click();", backButton);
////
////            System.out.println("Back button clicked twice successfully.");
////            Thread.sleep(3000);
////        } catch (Exception e) {
////            System.out.println("❌ Exception in wmmobile(): " + e.getMessage());
////            e.printStackTrace();
////        }
////
////    }
//}





package MA_MSG_Suite_OB;

import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;
import com.google.gson.stream.JsonReader;
import io.github.bonigarcia.wdm.WebDriverManager;
import okhttp3.*;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.StringReader;
import java.text.DecimalFormat;
import java.time.Duration;
import java.util.*;
import java.util.stream.Collectors;

/**
 * Process:
 *  1) Fetch all OLPNs from Excel (sheet: Tasks, columns: Testcase, OLPNs)
 *  2) Fetch OLPNs' ShipViaId from API (single batch IN(...))
 *  3) Segregate OLPNs by ShipViaId
 *  4) Login
 *  5) Menu -> search "WM Mobile"
 *  6) In WM Mobile, search and open "JD Palletize oLPN PCCC"
 *  7) For each ShipViaId group:
 *     a) Create a NEW pallet id and enter it
 *     b) Enter OLPNs (one by one)
 *     c) Click End Pallet
 *  8) Repeat until all ShipVia groups are completed
 *
 * Depends on: ExcelReader, Main1_URL_Login1 (your existing helpers)
 */
public class Main9_Palletisation {

    public static WebDriver driver;
    public static int time =60;
    // Inline selectors (no external selector helper classes)
    private static final String END_PALLET_COMPONENT_ID = "action_endpallet_button";
    private static final String PALLET_INPUT_COMPONENT_ID = "acceptpallet_barcodetextfield_pallet";
    private static final String OLPN_INPUT_COMPONENT_ID = "acceptolpn_barcodetextfield_olpn";
    private static final String JD_PALLETIZE_LABEL_COMPONENT = "jdpalletizeolpnpccc";
    private static final String CONFIRM_YES_COMPONENT_ID = "confirm-yes"; // update if different in your dialog

    private static final OkHttpClient httpClient = new OkHttpClient.Builder()
            .followRedirects(true).followSslRedirects(true)
            .callTimeout(Duration.ofSeconds(time))
            .connectTimeout(Duration.ofSeconds(time))
            .readTimeout(Duration.ofSeconds(time))
            .writeTimeout(Duration.ofSeconds(time))
            .build();

    /**
     * Non-standard entry point so you can call: Main9_Palletisation.main(filePath, tc, env)
     */
    public static void main(String filePath, String testcase, String env) throws Exception {

        // 1) Fetch all OLPNs from Excel
        List<String> olpns = getOlpnsForTestcase(filePath, testcase);
        System.out.println("==================================================");
        System.out.println("📦 OLPNs fetched for testcase '" + testcase + "': " + olpns);
        System.out.println("==================================================");

        // 2) Get bearer token via Excel config (your helper pattern)
        String token = getAuthTokenFromExcel();

        // 2.5) Batch API: fetch ShipViaId for all OLPNs and group by ShipViaId
        Map<String, List<String>> byShipVia = groupOlpnsByShipViaIdBatch(olpns, token);

        System.out.println("==================================================");
        System.out.println("🔗 OLPN → ShipViaId mapping:");
        byShipVia.forEach((shipVia, list) -> list.forEach(id ->
                System.out.println(" OLPN: " + id + " → ShipViaId: " + shipVia)));
        System.out.println("==================================================");
        System.out.println("🗂️ Segregated OLPNs by ShipViaId:");
        byShipVia.forEach((shipVia, list) -> System.out.println(" " + shipVia + " -> " + list));
        System.out.println("==================================================");

        // 3) Login (uses your helper)
        WebDriverManager.chromedriver().setup();
        ChromeOptions options = new ChromeOptions().addArguments("--start-maximized");
        driver = new ChromeDriver(options);
        driver.manage().window().maximize();

        Main1_URL_Login1 login1 = new Main1_URL_Login1(driver, env);
        login1.execute();
        System.out.println("✅ Login done.");

        // 4) Menu -> "WM Mobile"
        SearchMenuWM("WM Mobile", "WMMobile");

        // 5) In WM Mobile -> open "JD Palletize oLPN PCCC"
        SearchInWmMobileByTransaction("JD Palletize oLPN PCCC");

        // 6/7/8) For each ShipVia group: NEW pallet -> enter OLPNs -> End Pallet
        for (Map.Entry<String, List<String>> entry : byShipVia.entrySet()) {
            String shipViaId = entry.getKey();
            List<String> groupOlpns = entry.getValue();

            // Ensure we are on the JD Palletize screen
            ensureJdPalletizeScreen();

            // a) Create and enter NEW pallet id (random 10-digit, first digit non-zero)
            String palletId = enterRandomScannedPalletAndSubmit();
            System.out.println("=== 🚚 ShipViaId: " + shipViaId + " | PalletId: " + palletId + " ===");

            // b) Enter OLPNs one-by-one for this shipVia
            enterShipViaOlpnsSequentially(groupOlpns);

            // c) End Pallet
            clickEndPallet();

            try {
                writePalletIdForGroup(filePath, testcase, groupOlpns, palletId);
            } catch (Exception ex) {
                System.err.println("⚠️ Failed to write palletId to Excel for ShipVia " + shipViaId +
                        ": " + ex.getClass().getSimpleName() + " - " + ex.getMessage());
            }
            // Small pause between groups
            Thread.sleep(1000);



            // Small pause between groups
            Thread.sleep(1000);
        }

        // driver.quit(); // uncomment if you want to close automatically
    }

    // -------------------------------------------------------------------------
    // Excel: Fetch OLPNs for a given Testcase from "Tasks" sheet
    // -------------------------------------------------------------------------
    public static List<String> getOlpnsForTestcase(String filePath, String testcase) {
        List<String> olpns = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(filePath);
             org.apache.poi.ss.usermodel.Workbook workbook = new org.apache.poi.xssf.usermodel.XSSFWorkbook(fis)) {

            org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheet("Tasks");
            if (sheet == null) {
                System.err.println("❌ Sheet 'Tasks' not found in: " + filePath);
                return olpns;
            }
            org.apache.poi.ss.usermodel.Row headerRow = sheet.getRow(0);
            if (headerRow == null) {
                System.err.println("❌ Header row missing in 'Tasks'.");
                return olpns;
            }

            int colTestcase = -1, colOlpns = -1;
            for (int c = headerRow.getFirstCellNum(); c <= headerRow.getLastCellNum(); c++) {
                org.apache.poi.ss.usermodel.Cell cell = headerRow.getCell(c, org.apache.poi.ss.usermodel.Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                String header = getCellString(cell);
                if (header == null) continue;
                if (header.trim().equalsIgnoreCase("Testcase")) colTestcase = c;
                else if (header.trim().equalsIgnoreCase("OLPNs")) colOlpns = c;
            }

            if (colTestcase == -1 || colOlpns == -1) {
                System.err.println("❌ Required columns not found. Testcase idx=" + colTestcase + ", OLPNs idx=" + colOlpns);
                return olpns;
            }

            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                org.apache.poi.ss.usermodel.Row row = sheet.getRow(r);
                if (row == null) continue;
                String tcValue = getCellString(row.getCell(colTestcase, org.apache.poi.ss.usermodel.Row.MissingCellPolicy.RETURN_BLANK_AS_NULL));
                if (tcValue == null) continue;
                if (tcValue.trim().equalsIgnoreCase(testcase.trim())) {
                    String olpnsCell = getCellString(row.getCell(colOlpns, org.apache.poi.ss.usermodel.Row.MissingCellPolicy.RETURN_BLANK_AS_NULL));
                    if (olpnsCell != null && !olpnsCell.trim().isEmpty()) {
                        String[] parts = olpnsCell.split("[,\\n]");
                        for (String p : parts) {
                            String t = p.trim();
                            if (!t.isEmpty()) olpns.add(t);
                        }
                    }
                }
            }
        } catch (IOException e) {
            System.err.println("❌ IO Error reading Excel: " + e.getMessage());
        } catch (Exception e) {
            System.err.println("❌ Unexpected error: " + e.getMessage());
        }
        return olpns;
    }

    private static String getCellString(org.apache.poi.ss.usermodel.Cell cell) {
        if (cell == null) return null;
        switch (cell.getCellType()) {
            case STRING: {
                String s = cell.getStringCellValue();
                return (s == null) ? null : s.trim();
            }
            case NUMERIC:
                return new DecimalFormat("#.################").format(cell.getNumericCellValue()).trim();
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue()).trim();
            case FORMULA:
                org.apache.poi.ss.usermodel.CellType cached = cell.getCachedFormulaResultType();
                if (cached == org.apache.poi.ss.usermodel.CellType.STRING) {
                    String fs = cell.getStringCellValue();
                    return (fs == null) ? null : fs.trim();
                } else if (cached == org.apache.poi.ss.usermodel.CellType.NUMERIC) {
                    return new DecimalFormat("#.################").format(cell.getNumericCellValue()).trim();
                } else if (cached == org.apache.poi.ss.usermodel.CellType.BOOLEAN) {
                    return String.valueOf(cell.getBooleanCellValue()).trim();
                } else return null;
            default:
                return null;
        }
    }

    // -------------------------------------------------------------------------
    // API: Batch call with IN(olpn1,olpn2,...) and group by ShipViaId
    // -------------------------------------------------------------------------
    public static Map<String, List<String>> groupOlpnsByShipViaIdBatch(List<String> olpns, String bearerToken) throws IOException {
        Map<String, List<String>> grouped = new LinkedHashMap<>();
        if (olpns == null || olpns.isEmpty()) {
            System.err.println("⚠️ No OLPNs provided for grouping.");
            return grouped;
        }

        String inList = olpns.stream().filter(s -> s != null && !s.isBlank()).map(String::trim).collect(Collectors.joining(","));
        String body = "{\"Query\":\"OlpnId in (" + inList + ")\"}";
        System.out.println("📤 Batch body: " + body);

        ExcelReader reader = new ExcelReader();
        String BASE_URL = reader.getCellValueByHeader(1, "BASE_URL");
        String SelectedOrganization = reader.getCellValueByHeader(1, "SelectedOrganization");
        String SelectedLocation = reader.getCellValueByHeader(1, "SelectedLocation");
        reader.close();

        RequestBody requestBody = RequestBody.create(MediaType.parse("application/json"), body);
        Request.Builder rb = new Request.Builder()
                .url(BASE_URL + "/pickpack/api/pickpack/olpn/search")
                .post(requestBody)
                .addHeader("Content-Type", "application/json")
                .addHeader("Accept", "application/json")
                .addHeader("SelectedOrganization", SelectedOrganization)
                .addHeader("SelectedLocation", SelectedLocation)
                .addHeader("ComponentName", "com-manh-cp-pickpack")
                .addHeader("X-Requested-With", "XMLHttpRequest");
        if (bearerToken != null && !bearerToken.isBlank())
            rb.addHeader("Authorization", "Bearer " + bearerToken.trim());

        Request request = rb.build();
        Map<String, String> olpnShipViaMap = new HashMap<>();
        try (Response response = httpClient.newCall(request).execute()) {
            String responseBody = response.body() != null ? response.body().string() : "";
            String contentType = Optional.ofNullable(response.header("Content-Type")).orElse("");

            if (response.code() >= 200 && response.code() < 300 && contentType.toLowerCase(Locale.ROOT).contains("application/json")) {
                JsonReader jr = new JsonReader(new StringReader(responseBody));
                jr.setLenient(true);
                JsonElement rootEl = JsonParser.parseReader(jr);
                if (rootEl.isJsonObject()) {
                    JsonObject root = rootEl.getAsJsonObject();
                    if (root.has("data") && root.get("data").isJsonArray()) {
                        JsonArray data = root.get("data").getAsJsonArray();
                        for (JsonElement el : data) {
                            if (!el.isJsonObject()) continue;
                            JsonObject item = el.getAsJsonObject();
                            String olpn = readString(item, "OlpnId");
                            if (olpn == null || olpn.isBlank()) continue;
                            String shipVia = readString(item, "ShipViaId");
                            if (isBlank(shipVia)) shipVia = readString(item, "FinalDeliveryShipViaId");
                            if (isBlank(shipVia)) shipVia = readString(item, "ServiceLevelId");
                            if (isBlank(shipVia)) shipVia = readString(item, "StaticRouteId");
                            if (isBlank(shipVia)) {
                                for (Map.Entry<String, JsonElement> e : item.entrySet()) {
                                    String key = e.getKey();
                                    if (key.toLowerCase(Locale.ROOT).contains("shipviaid") && !e.getValue().isJsonNull()) {
                                        String val = e.getValue().getAsString();
                                        if (!isBlank(val)) {
                                            shipVia = val.trim();
                                            break;
                                        }
                                    }
                                }
                            }
                            if (isBlank(shipVia)) shipVia = "UNKNOWN";
                            olpnShipViaMap.put(olpn, shipVia);
                        }
                    }
                }
            }
        } catch (Exception e) {
            System.err.println("❌ Batch request failed: " + e);
        }

        // Group OLPNs by ShipViaId
        for (String id : olpns) {
            String shipVia = olpnShipViaMap.getOrDefault(id, "UNKNOWN");
            grouped.computeIfAbsent(shipVia, k -> new ArrayList<>()).add(id);
            System.out.println("↪️ OLPN " + id + " → ShipViaId: " + shipVia);
        }
        return grouped;
    }

    private static boolean isBlank(String s) {
        return s == null || s.trim().isEmpty();
    }

    private static String readString(JsonObject obj, String key) {
        return (obj.has(key) && !obj.get(key).isJsonNull()) ? obj.get(key).getAsString().trim() : null;
    }

    // -------------------------------------------------------------------------
    // Menu & Transaction open helpers (from your pallet.txt, tidied slightly)
    // -------------------------------------------------------------------------

    public static void SearchMenuWM(String Keyword, String id) {
        WebDriverWait wait1 = new WebDriverWait(driver, Duration.ofSeconds(time));
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(time));
        JavascriptExecutor js = (JavascriptExecutor) driver;

        int maxRetries = 6;
        for (int attempt = 1; attempt <= maxRetries; attempt++) {
            System.out.println("⏳ Waiting 10 seconds before refreshing...");
            try { Thread.sleep(10000); } catch (InterruptedException e) { throw new RuntimeException(e); }

            try {
                WebElement refreshBtn = wait.until(
                        ExpectedConditions.elementToBeClickable(By.cssSelector("ion-button[data-component-id='refresh']"))
                );
                refreshBtn.click();
                System.out.println("🔄 Refresh button clicked.");
                Thread.sleep(3000);
            } catch (Exception e) {
                System.err.println("⚠️ Unable to click refresh: " + e.getMessage());
            }

            try {
                WebElement shadowHost = wait1.until(ExpectedConditions.presenceOfElementLocated(
                        By.cssSelector("ion-button[data-component-id='menu-toggle-button']")
                ));
                SearchContext shadowRoot = (SearchContext) js.executeScript("return arguments[0].shadowRoot", shadowHost);
                WebElement nativeButton = shadowRoot.findElement(By.cssSelector("button.button-native"));
                wait1.until(ExpectedConditions.invisibilityOfElementLocated(By.cssSelector("manh-overlay-container")));
                js.executeScript("arguments[0].click();", nativeButton);
                System.out.println("✅ Menu toggle button clicked.");
                break;
            } catch (Exception e) {
                System.err.println("Error opening menu: " + e.getMessage());
            }
        }

        try {
            WebElement innerInput = wait.until(ExpectedConditions.presenceOfElementLocated(
                    By.cssSelector("ion-input[data-component-id='search-input'] input.native-input")
            ));
            wait.until(ExpectedConditions.elementToBeClickable(innerInput));
            innerInput.clear();
            innerInput.sendKeys(Keyword);
            System.out.println("🔎 Search typed: " + Keyword);
        } catch (Exception e) {
            System.err.println("❌ Error interacting with search input: " + e.getMessage());
        }

        try {
            WebElement element = wait.until(ExpectedConditions.presenceOfElementLocated(
                    By.cssSelector("button#wmMobile[data-component-id=" + id + "]")
            ));
            js.executeScript("arguments[0].scrollIntoView({block: 'center'});", element);
            js.executeScript("arguments[0].click();", element);
            System.out.println("✅ Clicked element with id: " + id);
            ArrayList<String> tabs = new ArrayList<>(driver.getWindowHandles());
            if (tabs.size() > 1) {
                driver.switchTo().window(tabs.get(1));
            }
        } catch (Exception e) {
            System.err.println("❌ Failed to click element with id: " + id);
            e.printStackTrace();
        }
    }


    static void SearchInWmMobileByTransaction(String transactionText) {
        JavascriptExecutor js = (JavascriptExecutor) driver;
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(time));

        Map<String, String> txnToComponentId = new HashMap<>();
        txnToComponentId.put("JD Palletize oLPN PCCC", JD_PALLETIZE_LABEL_COMPONENT);
        String componentId = txnToComponentId.getOrDefault(transactionText, null);

        try {
            WebElement searchInput = wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("ion-searchbar[data-component-id='search'] input[type='search']")));
            searchInput.clear();
            searchInput.sendKeys(transactionText);
            searchInput.clear();
            searchInput.sendKeys(transactionText);
            searchInput.sendKeys(Keys.ENTER);
        } catch (Exception e) {
            System.err.println("❌ WM search error: " + e.getMessage());
        }

        boolean clicked = false;
        if (componentId != null && !componentId.isBlank()) {
            try {
                WebElement labelElement = wait.until(ExpectedConditions.presenceOfElementLocated(
                        By.cssSelector("ion-label[data-component-id='" + componentId + "']")
                ));
                js.executeScript("arguments[0].scrollIntoView({block:'center'});", labelElement);
                wait.until(ExpectedConditions.elementToBeClickable(labelElement));
                js.executeScript("arguments[0].click();", labelElement);
                clicked = true;
            } catch (Exception ignored) {
            }
        }
        if (!clicked) {
            try {
                By byText = By.xpath("//*[normalize-space(text())=" + escapeForXPath(transactionText) + "]");
                WebElement labelByText = wait.until(ExpectedConditions.elementToBeClickable(byText));
                js.executeScript("arguments[0].scrollIntoView(true);", labelByText);
                js.executeScript("arguments[0].click();", labelByText);
            } catch (Exception ex) {
                System.err.println("❌ Fallback click failed: " + ex.getMessage());
            }
        }
    }

    private static String escapeForXPath(String s) {
        if (s == null) return "\"\"";
        if (s.indexOf('"') == -1) return "\"" + s + "\"";
        String[] parts = s.split("\"");
        StringBuilder sb = new StringBuilder("concat(");
        for (int i = 0; i < parts.length; i++) {
            sb.append("\"").append(parts[i]).append("\"");
            if (i < parts.length - 1) sb.append(", '\"', ");
        }
        sb.append(")");
        return sb.toString();
    }

    // -------------------------------------------------------------------------
    // JD Palletize: Pallet & OLPN entry (inline selectors; from your working methods)
    // -------------------------------------------------------------------------
    static String enterRandomScannedPalletAndSubmit() {
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(time));
        JavascriptExecutor js = (JavascriptExecutor) driver;
        java.security.SecureRandom rnd = new java.security.SecureRandom();
        StringBuilder sb = new StringBuilder(10);
        sb.append(1 + rnd.nextInt(9));
        for (int i = 1; i < 10; i++) sb.append(rnd.nextInt(10));
        String palletId = sb.toString();

        try {
            WebElement input = wait.until(ExpectedConditions.presenceOfElementLocated(
                    By.cssSelector("input[data-component-id='" + PALLET_INPUT_COMPONENT_ID + "']")
            ));
            wait.until(ExpectedConditions.elementToBeClickable(input));
            js.executeScript("arguments[0].scrollIntoView({block:'center'});", input);
            try {
                input.click();
            } catch (Exception e) {
                js.executeScript("arguments[0].click();", input);
            }
            try {
                input.clear();
            } catch (Exception ignored) {
                js.executeScript("arguments[0].value='';", input);
            }
            js.executeScript("arguments[0].focus();", input);
            input.sendKeys(palletId);
            input.sendKeys(Keys.ENTER);
            System.out.println("🧺 PalletId entered: " + palletId);
        } catch (Exception e) {
            System.err.println("❌ Pallet input failed: " + e.getMessage());
            // Fallback by placeholder (if your app uses a placeholder attr)
            try {
                By fallback = By.xpath("//input[@placeholder='Scanned Pallet' and contains(@class,'input-field')]");
                WebElement inputFallback = wait.until(ExpectedConditions.elementToBeClickable(fallback));
                js.executeScript("arguments[0].scrollIntoView({block:'center'});", inputFallback);
                js.executeScript("arguments[0].click();", inputFallback);
                inputFallback.sendKeys(palletId);
                inputFallback.sendKeys(Keys.ENTER);
            } catch (Exception ex) {
                System.err.println("❌ Pallet fallback failed: " + ex.getMessage());
            }
        }
        return palletId;
    }

    static void enterShipViaOlpnsSequentially(List<String> shipViaOlpns) {
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(time));
        JavascriptExecutor js = (JavascriptExecutor) driver;

        WebElement input;
        try {
            input = wait.until(ExpectedConditions.presenceOfElementLocated(
                    By.cssSelector("input[data-component-id='" + OLPN_INPUT_COMPONENT_ID + "']")
            ));
        } catch (Exception e) {
            input = wait.until(ExpectedConditions.presenceOfElementLocated(
                    By.xpath("//input[@placeholder='ScannedOlpn' and contains(@class,'input-field')]")
            ));
        }

        for (String olpn : shipViaOlpns) {
            if (olpn == null || olpn.trim().isEmpty()) {
                System.err.println("⚠️ Skipping blank OLPN.");
                continue;
            }
            String value = olpn.trim();
            try {
                try {
                    wait.until(ExpectedConditions.visibilityOf(input));
                    wait.until(ExpectedConditions.elementToBeClickable(input));
                } catch (StaleElementReferenceException sere) {
                    try {
                        input = wait.until(ExpectedConditions.presenceOfElementLocated(
                                By.cssSelector("input[data-component-id='" + OLPN_INPUT_COMPONENT_ID + "']")
                        ));
                    } catch (Exception e) {
                        input = wait.until(ExpectedConditions.presenceOfElementLocated(
                                By.xpath("//input[@placeholder='ScannedOlpn' and contains(@class,'input-field')]")
                        ));
                    }
                }
                js.executeScript("arguments[0].scrollIntoView({block:'center'});", input);
                try {
                    input.click();
                } catch (Exception clickIntercept) {
                    js.executeScript("arguments[0].click();", input);
                }
                js.executeScript("arguments[0].focus();", input);
                try {
                    input.clear();
                } catch (Exception ignored) {
                    js.executeScript("arguments[0].value='';", input);
                }
                input.sendKeys(value);
                input.sendKeys(Keys.ENTER);
                System.out.println("➕ OLPN added: " + value);
                try {
                    Thread.sleep(300);
                } catch (InterruptedException ie) {
                    Thread.currentThread().interrupt();
                }
            } catch (Exception e) {
                System.err.println("❌ Failed OLPN '" + value + "': " + e.getMessage());
                try {
                    input = wait.until(ExpectedConditions.elementToBeClickable(
                            By.xpath("//input[@placeholder='ScannedOlpn' and contains(@class,'input-field')]")
                    ));
                    js.executeScript("arguments[0].scrollIntoView({block:'center'});", input);
                    js.executeScript("arguments[0].click();", input);
                    js.executeScript("arguments[0].focus();", input);
                    input.clear();
                    input.sendKeys(value);
                    input.sendKeys(Keys.ENTER);
                    System.out.println("✅ Fallback OLPN: " + value);
                    try {
                        Thread.sleep(300);
                    } catch (InterruptedException ie) {
                        Thread.currentThread().interrupt();
                    }
                } catch (Exception ex) {
                    System.err.println("❌ Fallback also failed for OLPN '" + value + "': " + ex.getMessage());
                }
            }
        }
    }

    // -------------------------------------------------------------------------
    // End Pallet (inline selectors; tries ion-button and plain button)
    // -------------------------------------------------------------------------
    private static void clickEndPallet() {
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(time));
        JavascriptExecutor js = (JavascriptExecutor) driver;

        WebElement btn = null;
        try {
            btn = wait.until(ExpectedConditions.elementToBeClickable(
                    By.cssSelector("ion-button[data-component-id='" + END_PALLET_COMPONENT_ID + "']")
            ));
        } catch (Exception ignored) {
        }
        if (btn == null) {
            try {
                btn = wait.until(ExpectedConditions.elementToBeClickable(
                        By.cssSelector("button[data-component-id='" + END_PALLET_COMPONENT_ID + "']")
                ));
            } catch (Exception ignored) {
            }
        }
        if (btn == null) {
            System.err.println("❌ End Pallet button not found. Check visibility or component-id.");
            return;
        }

        js.executeScript("arguments[0].scrollIntoView({block:'center'});", btn);
        js.executeScript("arguments[0].click();", btn);
        System.out.println("🏁 End Pallet clicked (component-id: " + END_PALLET_COMPONENT_ID + ")");

        // Optional confirm dialog
        try {
            WebElement confirmYes = wait.until(ExpectedConditions.elementToBeClickable(
                    By.cssSelector("ion-button[data-component-id='" + CONFIRM_YES_COMPONENT_ID + "']")
            ));
            js.executeScript("arguments[0].click();", confirmYes);
            System.out.println("✅ End Pallet confirmed.");
        } catch (Exception ignored) {
        }
    }

    // -------------------------------------------------------------------------
    // Screen safety: ensure JD Palletize screen is active
    // -------------------------------------------------------------------------
    private static void ensureJdPalletizeScreen() {
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(time));
        try {
            wait.until(ExpectedConditions.presenceOfElementLocated(
                    By.cssSelector("ion-label[data-component-id='" + JD_PALLETIZE_LABEL_COMPONENT + "']")
            ));
        } catch (Exception ignored) {
            SearchInWmMobileByTransaction("JD Palletize oLPN PCCC");
        }
    }

    // -------------------------------------------------------------------------
    // Token retrieval from Excel (your existing pattern)
    // -------------------------------------------------------------------------
    public static String getAuthTokenFromExcel() throws IOException {
        ExcelReader reader = new ExcelReader();
        String LOGIN_URL = reader.getCellValueByHeader(1, "LOGIN_URL");
        String UIUsername = reader.getCellValueByHeader(1, "username");
        String UIPassword = reader.getCellValueByHeader(1, "password");
        reader.close();

        OkHttpClient client = new OkHttpClient();
        RequestBody body = RequestBody.create(MediaType.parse("application/x-www-form-urlencoded"),
                "grant_type=password&username=" + UIUsername + "&password=" + UIPassword);
        Request request = new Request.Builder()
                .url(LOGIN_URL)
                .post(body)
                .addHeader("Content-Type", "application/x-www-form-urlencoded")
                .addHeader("Authorization", "Basic dWpkc3N0YWdlMTpFYXJ0aC1Nb29uLVN1bjE=") // same as your file
                .build();
        Response response = client.newCall(request).execute();
        String responseBody = response.body() != null ? response.body().string() : null;
        JsonObject json = JsonParser.parseString(responseBody).getAsJsonObject();
        return json.has("access_token") ? json.get("access_token").getAsString() : null;
    }



    //     (F) Excel: write pallet id per OLPN group into column G ("pallet")
    public static void writePalletIdForGroup(String filePath,
                                             String testcase,
                                             List<String> groupOlpns,
                                             String palletId) {
        try (java.io.FileInputStream fis = new java.io.FileInputStream(filePath);


             org.apache.poi.ss.usermodel.Workbook wb = new org.apache.poi.xssf.usermodel.XSSFWorkbook(fis)) {

            org.apache.poi.ss.usermodel.Sheet sheet = wb.getSheet("Tasks");
            if (sheet == null) { System.err.println("❌ Sheet 'Tasks' not found: " + filePath); return; }

            org.apache.poi.ss.usermodel.Row header = sheet.getRow(0);
            if (header == null) { System.err.println("❌ Header row missing in 'Tasks'."); return; }

            int colTestcase = findCol(header, "Testcase");
            int colOlpns = findCol(header, "OLPNs");
            int colPallet = findOrCreateCol(header, "pallet", 6); // 0-based → G

            if (colTestcase == -1 || colOlpns == -1) {
                System.err.println("❌ Required columns missing. Testcase=" + colTestcase + ", OLPNs=" + colOlpns);
                return;
            }

            java.util.Set<String> groupSet = new java.util.HashSet<>(
                    groupOlpns.stream().map(String::trim).collect(java.util.stream.Collectors.toSet())
            );

            // Walk all rows; write palletId for rows that match testcase and contain any OLPN from the group
            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                org.apache.poi.ss.usermodel.Row row = sheet.getRow(r);
                if (row == null) continue;

                String tc = getCellString(row.getCell(colTestcase, org.apache.poi.ss.usermodel.Row.MissingCellPolicy.RETURN_BLANK_AS_NULL));
                if (tc == null || !tc.trim().equalsIgnoreCase(testcase.trim())) continue;

                String olpnsCell = getCellString(row.getCell(colOlpns, org.apache.poi.ss.usermodel.Row.MissingCellPolicy.RETURN_BLANK_AS_NULL));
                if (olpnsCell == null || olpnsCell.isEmpty()) continue;

                java.util.List<String> tokens = splitOlpns(olpnsCell);
                boolean hasAny = tokens.stream().anyMatch(t -> groupSet.contains(t));

                if (hasAny) {
                    org.apache.poi.ss.usermodel.Cell palletCell =
                            row.getCell(colPallet, org.apache.poi.ss.usermodel.Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    palletCell.setCellValue(palletId);
                }
            }

            // Save workbook back to disk
            try (java.io.FileOutputStream fos = new java.io.FileOutputStream(filePath)) {
                wb.write(fos);
            }
            System.out.println("📝 PalletId '" + palletId + "' written to column G for matching rows.");
        } catch (Exception ex) {
            System.err.println("❌ Failed to write pallet id: " + ex.getClass().getSimpleName() + " - " + ex.getMessage());
        }
    }






    /** Find a header column by name (case-insensitive). */
    private static int findCol(org.apache.poi.ss.usermodel.Row headerRow, String name) {
        for (int c = headerRow.getFirstCellNum(); c <= headerRow.getLastCellNum(); c++) {
            String h = getCellString(headerRow.getCell(c, org.apache.poi.ss.usermodel.Row.MissingCellPolicy.RETURN_BLANK_AS_NULL));
            if (h != null && h.trim().equalsIgnoreCase(name)) return c;
        }
        return -1;
    }

    /** Find header 'name' or create it at preferredIndex (0-based). */
    private static int findOrCreateCol(org.apache.poi.ss.usermodel.Row headerRow, String name, int preferredIndex) {
        int idx = findCol(headerRow, name);
        if (idx != -1) return idx;
        org.apache.poi.ss.usermodel.Cell cell = headerRow.getCell(preferredIndex, org.apache.poi.ss.usermodel.Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
        cell.setCellValue(name);
        return preferredIndex;
    }

    /** Robust split for OLPNs stored as comma and/or newline separated values. */
    private static java.util.List<String> splitOlpns(String cellValue) {
        return java.util.Arrays.stream(cellValue.split("[,\\n]+"))
                .map(String::trim)
                .filter(s -> !s.isEmpty())
                .collect(java.util.stream.Collectors.toList());
    }















}