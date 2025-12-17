package MA_MSG_Suite_OB;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.FileInputStream;
import java.time.Duration;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
//import java.util.*;

public class FinalPacking {
    public static WebDriver driver;
    public static String filePath = "C:\\Users\\2210420\\IdeaProjects\\Testcases\\OOdata.xlsx";
    // public static String StaticRoute = "ALICANTE-HUB-01";//Load
    public static String StaticRoute;//Load
    public static String location;
    //public static String location = "CONHUBALICANTE";
    public static String shipmentNumber;
    public static String olpn = "KU700311001557";

    public static String Carrier = "USPS";//Load
    public static String Trailer = "TRL_VAB012";//Load
    public static String Trailertype = "JD40SDT";//LoaD
    public static String DockDoor = "DD03";//Load
    //  public static String pallet = "PLT0203501"; //Load + close shipment-10digit
    public static String pallet; //Load + close shipment-10digit
    public static String sealinput = "90000008"; //close shipment

    public static void main(String[] args) throws InterruptedException {
        login();
        //  Fetching_data_from_Excel(filePath);f
        SearchandOpenWMMobie();
//        SearchInWmMobile("JD Pack Pick Cart", "jdpackpickcart");
//        Packing(filePath);
//        SearchInWmMobile( "JD Palletize oLPN PCCC", "jdpalletizeolpnpccc");
//        Palletise(filePath);
        SearchInWmMobile( "JD OB Putaway To Staging", "jdobputawaytostaging");
        OutboundPutaway(filePath);
//
//        Shipment();
//        CheckIn();
//       Load();
//       JDCloseShipment();

    }

    public static void login() {
        WebDriverManager.chromedriver().setup();
        driver = new ChromeDriver();
        driver.manage().window().maximize();

        driver.get("https://ujdss-auth.sce.manh.com/auth/realms/maactive/protocol/openid-connect/auth?scope=openid&client_id=zuulserver.1.0.0&redirect_uri=https://ujdss.sce.manh.com/login&response_type=code&state=52FrgC");

        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(20));
        WebElement usernameField = wait.until(ExpectedConditions.elementToBeClickable(By.id("username")));
        WebElement passwordField = driver.findElement(By.id("password"));

        usernameField.sendKeys("cogs");
        passwordField.sendKeys("Cogs@123456");

        driver.findElement(By.id("kc-login")).click();
        wait.until(ExpectedConditions.urlContains("ujdss.sce.manh.com/udc/landing"));
    }

    static void SearchandOpenWMMobie() {
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
        JavascriptExecutor js = (JavascriptExecutor) driver;

        try {
            // 🔹 Step 1: Open the side menu
            WebElement shadowHost = driver.findElement(By.cssSelector("ion-button.menu-toggle-button"));
            SearchContext shadowRoot = (SearchContext) js.executeScript("return arguments[0].shadowRoot", shadowHost);
            WebElement menuButton = shadowRoot.findElement(By.cssSelector("button.button-native"));
            menuButton.click();

            // 🔹 Step 2: Search for "WM Mobile" in the menu
            WebElement searchInput = wait.until(ExpectedConditions.elementToBeClickable(
                    By.xpath("//input[@placeholder='Search Menu...']")));
            searchInput.clear();
            searchInput.sendKeys("WM Mobile");
//
            try {
                // Wait for the button to be present and visible
                WebElement wmMobileButton = wait.until(ExpectedConditions.presenceOfElementLocated(
                        By.cssSelector("button#wmMobile[data-component-id='WMMobile']")
                ));

                // Scroll into view and click using JS (in case normal click fails)
                js.executeScript("arguments[0].scrollIntoView(true);", wmMobileButton);
                js.executeScript("arguments[0].click();", wmMobileButton);
                System.out.println("✅ 'WM Mobile' button clicked.");
            } catch (Exception e) {
                System.err.println("❌ Error clicking 'WM Mobile': " + e.getMessage());
                e.printStackTrace();
            }
            // 🔹 Step 4: Switch to the new tab
            Thread.sleep(3000);
            ArrayList<String> tabs = new ArrayList<>(driver.getWindowHandles());
            driver.switchTo().window(tabs.get(1));
            Thread.sleep(3000);

        } catch (InterruptedException e) {
            throw new RuntimeException(e);
        }

    }
    static void SearchInWmMobile(String Transaction, String ComponentId)  {
        JavascriptExecutor js = (JavascriptExecutor) driver;
        // Wait for the ion-item with data-component-id 'jdmakepickcart'
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));


        try {
            WebElement searchInput1 = wait.until(ExpectedConditions.elementToBeClickable(
                    By.cssSelector("input.searchbar-input[placeholder='Search']")));
//                    By.xpath("//input[@type='search' and @placeholder='Search']")));
            searchInput1.click();
            searchInput1.clear();
            Thread.sleep(3000);
            searchInput1.sendKeys(Transaction);
        } catch (Exception e) {
            System.err.println("❌ Error in JD Make Pick Cart " + e.getMessage());
            e.printStackTrace();
        }
        // Locate the ion-label using its data-component-id
        WebElement labelElement = driver.findElement(
                By.cssSelector("ion-label[data-component-id='" + ComponentId + "']")
        );

        // Scroll into view to ensure it's interactable
        ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", labelElement);

        // Click using JavaScript (in case native click doesn't work)
        ((JavascriptExecutor) driver).executeScript("arguments[0].click();", labelElement);

        System.out.println("Clicked on 'JD Make Pick Cart' label.");


    }

    public static void Packing(String filePath) {
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(20));
        JavascriptExecutor js = (JavascriptExecutor) driver;

        try {


            try (FileInputStream fis = new FileInputStream(filePath);
                 Workbook workbook = new XSSFWorkbook(fis)) {

                Sheet sheet = workbook.getSheetAt(3); // first sheet

                // Map to group OLPNs by testcase
                Map<String, List<String>> testcaseMap = new LinkedHashMap<>();

                // Iterate rows (assuming first row is header)
                for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                    Row row = sheet.getRow(i);
                    if (row == null) continue;

                    String olpn = row.getCell(4).getStringCellValue();       // column E
                    String testcase = row.getCell(0).getStringCellValue();   // column A

                    testcaseMap.computeIfAbsent(testcase, k -> new ArrayList<>()).add(olpn);
                }

                workbook.close();
                fis.close();



                // Process OLPNs grouped by testcase
                for (Map.Entry<String, List<String>> entry : testcaseMap.entrySet()) {
                    String testcase = entry.getKey();
                    List<String> olpns = entry.getValue();

                    System.out.println("=== Processing Testcase: " + testcase + " ===");

                    for (String olpn : olpns) {
                        System.out.println("Packing OLPN: " + olpn);

                        // Click Skip Printer Scan button
                        WebElement skipPrinterScanButton = wait.until(
                                ExpectedConditions.elementToBeClickable(
                                        By.cssSelector("button[data-component-id='action_skipprinterscan_button']")
                                )
                        );
                        skipPrinterScanButton.click();

                        // Enter OLPN
                        WebElement olpnInput = wait.until(
                                ExpectedConditions.visibilityOfElementLocated(
                                        By.cssSelector("input[data-component-id='acceptolpn_barcodetextfield_olpn']")
                                )
                        );
                        olpnInput.clear();
                        olpnInput.sendKeys(olpn);
                        olpnInput.sendKeys(Keys.ENTER);

                        // Optional: add validation/wait between OLPNs

                        try {
                            wait = new WebDriverWait(driver, Duration.ofSeconds(10));
                            WebElement okButton = wait.until(ExpectedConditions.elementToBeClickable(
                                    By.xpath("//button[.//span[text()='Ok']]")
                            ));

                            ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", okButton);
                            ((JavascriptExecutor) driver).executeScript("arguments[0].click();", okButton);

                            System.out.println("Clicked on 'Ok' button in alert.");
                        } catch ( ElementClickInterceptedException ex) {
                            System.out.println("Alert not found or click failed. Please click manually.");
                            Thread.sleep(10000);
                            System.out.println("10 sec remaining");
                            Thread.sleep(10000);
                        }
                    }
                }

            }

            Thread.sleep(3000); // Or use WebDriverWait if preferred


            // Wait until the ion-buttons container is present

            WebElement backButton = wait.until(ExpectedConditions.presenceOfElementLocated(
                    By.cssSelector("ion-buttons[data-component-id='action_back_button']")
            ));

// Scroll into view to ensure visibility
            ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", backButton);

// Click the button twice using JavaScript

            // js.executeScript("arguments[0].click();", backButton);
            Thread.sleep(300); // brief pause between clicks
            js.executeScript("arguments[0].click();", backButton);

            System.out.println("Back button clicked twice successfully.");
            Thread.sleep(3000);
        } catch (Exception e) {
            System.out.println("❌ Exception in wmmobile(): " + e.getMessage());
            e.printStackTrace();
        }

    }



    public static void Palletise(String filePath1) throws InterruptedException{
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
        JavascriptExecutor js = (JavascriptExecutor) driver;

        try {
            try (FileInputStream fis1 = new FileInputStream(filePath1);
                 Workbook workbook = new XSSFWorkbook(fis1)) {

                Sheet sheet = workbook.getSheetAt(3); // first sheet

                // Map to group OLPNs by testcase
                Map<String, List<String>> testcaseMap = new LinkedHashMap<>();

                // Iterate rows (assuming first row is header)
                for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                    Row row = sheet.getRow(i);
                    if (row == null) continue;

                    // Column A = Testcase, Column E = OLPN
                    Cell testcaseCell = row.getCell(0);
                    Cell olpnCell = row.getCell(4);

                    if (testcaseCell == null || olpnCell == null) continue;

                    String testcase = testcaseCell.getStringCellValue().trim();
                    String olpn = olpnCell.getStringCellValue().trim();

                    testcaseMap.computeIfAbsent(testcase, k -> new ArrayList<>()).add(olpn);
                }

                workbook.close();
                fis1.close();


                // WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));

                // Process OLPNs grouped by testcase
                int palletCounter = 1; // generate pallet IDs or use from Excel if available

                for (Map.Entry<String, List<String>> entry : testcaseMap.entrySet()) {
                    String testcase = entry.getKey();
                    List<String> olpns = entry.getValue();

                    System.out.println("=== Processing Testcase: " + testcase + " ===");

                    if (olpns.isEmpty()) continue;

                    // Derive pallet root from first OLPN
                    String firstOlpn = olpns.get(0);
                    String last6 = firstOlpn.length() >= 6 ?
                            firstOlpn.substring(firstOlpn.length() - 5) :
                            String.format("%05d", 0);

                    int palletIndex = 1;


                    // Loop through OLPNs, 20 per pallet
                    for (int i = 0; i < olpns.size(); i++) {

                        // At the start of each batch of 20 OLPNs → enter pallet once
                        if (i % 20 == 0) {
                            String pallet = String.format("PLT%s%02d", last6, palletIndex++);

                            WebElement palletInput = wait.until(
                                    ExpectedConditions.visibilityOfElementLocated(
                                            By.cssSelector("input[data-component-id='acceptpallet_barcodetextfield_pallet']")
                                    )
                            );
                            palletInput.clear();
                            palletInput.sendKeys(pallet);
                            palletInput.sendKeys(Keys.ENTER);
                            System.out.println("Entered pallet value: " + pallet);
                        }

                        // Enter OLPN (20 times per pallet)
                        String olpn = olpns.get(i);
                        WebElement olpnInput = wait.until(
                                ExpectedConditions.visibilityOfElementLocated(
                                        By.cssSelector("input[data-component-id='acceptolpn_barcodetextfield_olpn']")
                                )
                        );
                        olpnInput.clear();
                        olpnInput.sendKeys(olpn);
                        olpnInput.sendKeys(Keys.ENTER);
                        System.out.println("Entered OLPN value: " + olpn);

                        WebDriverWait wait1 = new WebDriverWait(driver, Duration.ofSeconds(10));

                        try {
                            // Wait until the Confirm button is visible and clickable
                            WebElement confirmButton = wait1.until(
                                    ExpectedConditions.elementToBeClickable(
                                            By.xpath("//div[contains(@class,'alert-button-group')]//button[.//span[text()='Confirm']]")
                                    )
                            );

                            confirmButton.click();
                            System.out.println("Clicked on Confirm button.");
                        } catch (Exception e) {
                            System.out.println("Alert did not appear or Confirm button not found.");
                        }

                        // If pallet reached 20 OLPNs OR last OLPN in testcase → click End Pallet
                        if ((i + 1) % 20 == 0 || i == olpns.size() - 1) {
                            WebElement endPalletButton = wait.until(
                                    ExpectedConditions.elementToBeClickable(
                                            By.cssSelector("button[data-component-id='action_endpallet_button']")
                                    )
                            );
                            endPalletButton.click();
                            System.out.println("Clicked End Pallet button.");
                        }


                    }
                }


            }

        } catch (Exception e) {
            System.err.println("Error in Palletising': " + e.getMessage());
            e.printStackTrace();
        }



        // Wait until the ion-buttons container is present

        WebElement backButton = wait.until(ExpectedConditions.presenceOfElementLocated(
                By.cssSelector("ion-buttons[data-component-id='action_back_button']")
        ));

// Scroll into view to ensure visibility
        ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", backButton);

// Click the button twice using JavaScript

        //  js.executeScript("arguments[0].click();", backButton);
        Thread.sleep(2000); // brief pause between clicks
        js.executeScript("arguments[0].click();", backButton);

        System.out.println("Back button clicked twice successfully.");
        Thread.sleep(3000);

    }



    public static void OutboundPutaway(String filePath) throws InterruptedException {
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
        JavascriptExecutor js = (JavascriptExecutor) driver;

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheet("Outbound");

            // Get Row 1 (index 0 because POI is zero-based)
            Row row = sheet.getRow(0);

            // Map cells to variables
            String pallet = row.getCell(0).getStringCellValue();  // Row 1, Col 1
//            String StaticRoute = row.getCell(1).getStringCellValue();  // Row 1, Col 2
//            String location = row.getCell(2).getStringCellValue();  // Row 1, Col 3
//            //   int quantity = (int) row.getCell(3).getNumericCellValue(); // Row 1, Col 4

            // Close resources
            workbook.close();
            fis.close();
            System.out.println("Fetching data done ");
            //return pallet;
        }
        catch (Exception e) {
            System.err.println("❌ Error in Fetching data " + e.getMessage());
            e.printStackTrace();
        }
        //  return pallet;






//        try {
//            Thread.sleep(3000);
//            WebElement searchInput1 = wait.until(ExpectedConditions.elementToBeClickable(
//                    By.cssSelector("input.searchbar-input[placeholder='Search']")));
////                    By.xpath("//input[@type='search' and @placeholder='Search']")));
//            searchInput1.click();
//            searchInput1.clear();
//            Thread.sleep(3000);
//            searchInput1.sendKeys("JD OB Putaway To Staging");
//        } catch (Exception e) {
//            System.err.println("❌ Error in JD OB Putaway To Staging " + e.getMessage());
//            e.printStackTrace();
//        }
        // Wait for the page to load (replace with WebDriverWait for production use)
//        Thread.sleep(1000);
//
//        // Locate the ion-label using its data-component-id
//        WebElement labelElement = driver.findElement(
//                By.cssSelector("ion-label[data-component-id='jdobputawaytostaging']")
//        );
//
//        // Scroll into view to ensure it's interactable
//        ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", labelElement);
//
//        // Click using JavaScript (in case native click doesn't work)
//        ((JavascriptExecutor) driver).executeScript("arguments[0].click();", labelElement);
//
//        System.out.println("Clicked on 'JD OB Putaway To Staging' label.");


// Optional: wait briefly to ensure stability
        Thread.sleep(2000); // Or use WebDriverWait if preferred


        // Locate the input field using its data-component-id
        WebElement inputField = driver.findElement(
                By.cssSelector("input[data-component-id='acceptcontainer_barcodetextfield_scancontainer']")
        );

        // Clear any existing text and enter the pallet value
        inputField.clear();
        inputField.sendKeys(pallet);

        System.out.println("Entered pallet value into Scan Container field: " + pallet);


        // Wait until the input field is visible and interactable
// Try to locate and click the OK button
        try {
            inputField.sendKeys(Keys.ENTER);
//            WebElement goButton4 = driver.findElement(
//                    By.cssSelector("ion-button[data-component-id='acceptcontainer_barcodetextfield_go']")
//            );
//
//            // Scroll into view
//            ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", goButton4);
//
//            // Click using JavaScript
//            ((JavascriptExecutor) driver).executeScript("arguments[0].click();", goButton4);

            System.out.println("Clicked on go button.");
        } catch ( ElementClickInterceptedException e) {
            System.out.println("Not worked");
        }



        Thread.sleep(5000);

        // Try to find scan location input
        try {
            WebElement inputField10 = driver.findElement(
                    By.cssSelector("input[data-component-id='acceptlocation_barcodetextfield_scanlocation']")
            );
            inputField10.clear();
            inputField10.sendKeys(location);
            System.out.println("Entered scan location: " + location);
        } catch (Exception e) {
            // Handle alert and enter destination location instead
            try {
                WebElement okButton = driver.findElement(
                        By.xpath("//div[contains(@class,'alert-button-group')]//button[.//span[text()='Ok']]")
                );
                ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", okButton);
                ((JavascriptExecutor) driver).executeScript("arguments[0].click();", okButton);
                System.out.println("Clicked on 'Ok' button in alert.");
            } catch (ElementClickInterceptedException ex) {
                System.out.println("Pallet is invalid OR  Click manually");
                Thread.sleep(10000);
                System.out.println("10 sec remaining");
                Thread.sleep(10000);
            }
            Thread.sleep(4000);

            WebElement inputField11 = driver.findElement(
                    By.cssSelector("input[data-component-id='acceptlocation_barcodetextfield_destinationlocation']")
            );
            inputField11.clear();
            inputField11.sendKeys(location);
            System.out.println("Entered location value: " + location);
            Thread.sleep(5000);

            inputField11.sendKeys(Keys.ENTER);


        }



//
        try {
            // inputField11.sendKeys(Keys.ENTER);
//            // Locate the ion-button using its data-component-id
//            WebElement goButton5 = driver.findElement(
//                    By.cssSelector("ion-button[data-component-id='acceptlocation_barcodetextfield_go']")
//            );
//
//            // Scroll into view to ensure it's interactable
//            ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", goButton5);
//
//            // Click using JavaScript (in case native click doesn't work)
//            ((JavascriptExecutor) driver).executeScript("arguments[0].click();", goButton5);
//
//            System.out.println("Clicked on 'Go' button for location barcode.");



            Thread.sleep(3000); // brief pause between clicks
            // Wait until the ion-buttons container is present

            WebElement backButton = wait.until(ExpectedConditions.presenceOfElementLocated(
                    By.cssSelector("ion-buttons[data-component-id='action_back_button']")
            ));

// Scroll into view to ensure visibility
            ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", backButton);

// Click the button twice using JavaScript

            // js.executeScript("arguments[0].click();", backButton);

            js.executeScript("arguments[0].click();", backButton);

            System.out.println("Back button clicked successfully. 02");
        } catch (ElementClickInterceptedException e) {
            System.out.println("Go Not worked");
            System.out.println(" waiting for Back button clicked ");
            Thread.sleep(3000); // brief pause between clicks
            // Wait until the ion-buttons container is present

            WebElement backButton = wait.until(ExpectedConditions.presenceOfElementLocated(
                    By.cssSelector("ion-buttons[data-component-id='action_back_button']")
            ));



        }

        Thread.sleep(3000);


        //  public static void tabswitch()

        // Get all open window handles
        ArrayList<String> tabs = new ArrayList<>(driver.getWindowHandles());

// Switch to the first tab
        driver.switchTo().window(tabs.get(0));

// Optional: Try to bring the tab to the front using JavaScript
        // JavascriptExecutor js = (JavascriptExecutor) driver;
        js.executeScript("window.focus();");

        System.out.println("✅ Switched to the first tab and attempted to bring it to the front.");


    }



    public static void Shipment() throws InterruptedException {
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(15));
        JavascriptExecutor js = (JavascriptExecutor) driver;

        // ✅ Wait for overlay to disappear
        wait.until(ExpectedConditions.invisibilityOfElementLocated(By.cssSelector("manh-overlay-container")));

        // ✅ Access shadow DOM safely
        WebElement shadowHost = wait.until(ExpectedConditions.presenceOfElementLocated(By.cssSelector("ion-button.menu-toggle-button")));
        SearchContext shadowRoot = (SearchContext) js.executeScript("return arguments[0].shadowRoot", shadowHost);
        WebElement menuButton = shadowRoot.findElement(By.cssSelector("button.button-native"));
        js.executeScript("arguments[0].click();", menuButton);

        // ✅ Search for Static Route
        WebElement searchInput = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@placeholder='Search Menu...']")));
        searchInput.clear();
        searchInput.sendKeys("Static Route");

//    WebElement StaticRoute = wait.until(ExpectedConditions.elementToBeClickable(By.id("StaticRoute")));
//    js.executeScript("arguments[0].click();", StaticRoute);

        WebElement staticRouteButton = wait.until(ExpectedConditions.elementToBeClickable(
                By.cssSelector("button#staticRoute[data-component-id='StaticRoute']")
        ));
        staticRouteButton.click();


        // Wait for presence of the ion-icon element
        WebElement closeIcon = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("close-menu-button")));

        // Check if visible before clicking
        if (closeIcon.isDisplayed()) {
            js.executeScript("arguments[0].click();", closeIcon);
            System.out.println("Clicked the visible close icon.");
        } else {
            System.out.println("Close icon is present but not visible.");
        }


        Thread.sleep(3000);


        // ✅ Open filter panel via shadow DOM
        WebElement filterBtnHost = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("(//ion-button[contains(@class,\"toggle-button\")])[3]")));
        js.executeScript("arguments[0].shadowRoot.querySelector('.button-inner').click();", filterBtnHost);

        Thread.sleep(3000);

        // ✅ Chevron down for Static Route
        WebElement chevronDown = wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("ion-button[data-component-id='StaticRoute-StaticRoute-chevron-down']")));
        js.executeScript("arguments[0].scrollIntoView(true);", chevronDown);
        js.executeScript("arguments[0].click();", chevronDown);
        System.out.println("StaticRoute chevron-down button clicked.");

        Thread.sleep(2000);

        // ✅ Enter Static Route
        WebElement orderPlanningRunInputField = wait.until(
                ExpectedConditions.elementToBeClickable(By.xpath("//ion-input[@data-component-id='StaticRouteId']//input"))
        );

        //String StaticRoute = "ALICANTE-HUB-01";
        if (StaticRoute != null && !StaticRoute.isEmpty()) {
            orderPlanningRunInputField.click();
            orderPlanningRunInputField.clear();
            orderPlanningRunInputField.sendKeys(StaticRoute);
            orderPlanningRunInputField.sendKeys(Keys.ENTER);
        }

        System.out.println("StaticRoute entered: " + StaticRoute);

        Thread.sleep(5000); // Consider replacing with dynamic wait if possible

        // ✅ Refresh
        WebElement refreshHost = wait.until(ExpectedConditions.presenceOfElementLocated(By.cssSelector("ion-button.refresh-button")));
        WebElement refreshButton = (WebElement) js.executeScript("return arguments[0].shadowRoot.querySelector('button.button-native')", refreshHost);
        js.executeScript("arguments[0].click();", refreshButton);
        System.out.println("Refresh clicked.");
        Thread.sleep(500);

//Activation of Shipment
        // Retry logic
        int retryCount = 0;
        boolean seeMoreFound = false;

        while (retryCount < 3 && !seeMoreFound) {
            try {
                // Select record before clicking ACTIVATE
                try {
                    WebElement cardView = wait.until(ExpectedConditions.elementToBeClickable(
                            By.cssSelector("card-view[data-component-id='Card-View'] .card-row.primary[tabindex='0']")
                    ));
                    js.executeScript("arguments[0].scrollIntoView(true);", cardView);
                    cardView.click();
                    System.out.println("Card view record clicked.");
                } catch (StaleElementReferenceException e) {
                    WebElement cardViewRetry = wait.until(ExpectedConditions.elementToBeClickable(
                            By.cssSelector("card-view[data-component-id='Card-View'] .card-row.primary[tabindex='0']")
                    ));
                    js.executeScript("arguments[0].scrollIntoView(true);", cardViewRetry);
                    cardViewRetry.click();
                    System.out.println("Card view record clicked after retry.");
                }

                // Click ACTIVATE button
                WebElement activateHost = wait.until(ExpectedConditions.presenceOfElementLocated(
                        By.cssSelector("ion-button[data-component-id='footer-panel-action-Activate']")
                ));
                WebElement activateButton = (WebElement) js.executeScript(
                        "return arguments[0].shadowRoot.querySelector('button.button-native')", activateHost
                );

                Thread.sleep(300);
                if (activateButton.isDisplayed()) {
                    js.executeScript("arguments[0].click();", activateButton);
                    System.out.println("ACTIVATE button clicked.");
                } else {
                    System.out.println("ACTIVATE button is present but not visible.");
                }

                Thread.sleep(300);

                // Check for toast visibility
                try {
                    wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector("div.toast-grid")));

                    // Click seeMoreButton
                    WebElement seeMoreButton = wait.until(ExpectedConditions.elementToBeClickable(
                            By.cssSelector("div.toast-action button[data-component-id='toast-link']")
                    ));
                    seeMoreButton.click();
                    System.out.println("seeMoreButton button clicked.");

                    // Extract shipment number
                    WebDriverWait wait1 = new WebDriverWait(driver, Duration.ofSeconds(10));
                    WebElement tdElement = wait1.until(ExpectedConditions.presenceOfElementLocated(
                            By.xpath("//td[contains(text(),'Shipment SHI')]")
                    ));

                    String tdText = tdElement.getText();
                    //String shipmentNumber = "";
                    Matcher matcher = Pattern.compile("Shipment\\s+(SHI\\d+)").matcher(tdText);
                    if (matcher.find()) {
                        shipmentNumber = matcher.group(1);
                    }
                    System.out.println("Shipment Number: " + shipmentNumber);

                    seeMoreFound = true; // Success

                } catch (TimeoutException te) {
                    retryCount++;
                    System.out.println("Toast not found. Retrying by selecting record... (Attempt " + retryCount + ")");
                }

            } catch (Exception e) {
                retryCount++;
                System.out.println("Unexpected error on attempt " + retryCount + ": " + e.getMessage());
            }
        }

        if (!seeMoreFound) {
            System.out.println("GOT AN ERROR WHILE CREATING SHIPMENT");
            System.exit(1); // Exit the program
        }
        Thread.sleep(1000);


        // Wait for the ion-button to be present
        WebDriverWait wait4 = new WebDriverWait(driver, Duration.ofSeconds(10));
        WebElement okButtonHost = wait4.until(ExpectedConditions.presenceOfElementLocated(
                By.cssSelector("ion-button[data-component-id='ion-button-process-ok-button-5763']")
        ));

// Access the shadow root and find the native button

        WebElement nativeButton = (WebElement) js.executeScript(
                "return arguments[0].shadowRoot.querySelector('button.button-native')", okButtonHost
        );

// Scroll into view and click
        js.executeScript("arguments[0].scrollIntoView(true);", nativeButton);
        js.executeScript("arguments[0].click();", nativeButton);

        System.out.println("Clicked on OK button inside modal.");


    }
    public static void CheckIn() throws InterruptedException {
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(15));
        JavascriptExecutor js = (JavascriptExecutor) driver;

        // ✅ Wait for overlay to disappear
        wait.until(ExpectedConditions.invisibilityOfElementLocated(By.cssSelector("manh-overlay-container")));

        // ✅ Access shadow DOM safely
        WebElement shadowHost = wait.until(ExpectedConditions.presenceOfElementLocated(By.cssSelector("ion-button.menu-toggle-button")));
        SearchContext shadowRoot = (SearchContext) js.executeScript("return arguments[0].shadowRoot", shadowHost);
        WebElement menuButton = shadowRoot.findElement(By.cssSelector("button.button-native"));
        js.executeScript("arguments[0].click();", menuButton);

        // ✅ Search for Check In
        WebElement searchInput = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@placeholder='Search Menu...']")));
        searchInput.clear();
        searchInput.sendKeys("Check In");

//    WebElement StaticRoute = wait.until(ExpectedConditions.elementToBeClickable(By.id("StaticRoute")));
//    js.executeScript("arguments[0].click();", StaticRoute);

        WebElement CheckInButton = wait.until(ExpectedConditions.elementToBeClickable(
                By.cssSelector("button#yardCheckIn[data-component-id='CheckIn']")
        ));
        CheckInButton.click();

        // Wait for presence of the ion-icon element
        WebElement closeIcon = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("close-menu-button")));

        // Check if visible before clicking
        if (closeIcon.isDisplayed()) {
            js.executeScript("arguments[0].click();", closeIcon);
            System.out.println("Clicked the visible close icon.");
        } else {
            System.out.println("Close icon is present but not visible.");
        }


        // Wait for the span element to be clickable
        WebDriverWait wait2 = new WebDriverWait(driver, Duration.ofSeconds(3));
        WebElement checkInInfo = wait2.until(ExpectedConditions.elementToBeClickable(
                By.cssSelector("span.expand-header-text[data-component-id='Check-InInfo']")
        ));

// Click the element
        checkInInfo.click();
        System.out.println("Check-In Info clicked.");


        // Locate the ion-input wrapper using data-component-id
        WebElement carrierInputWrapper = driver.findElement(
                By.cssSelector("ion-input[data-component-id='CarrierId']")
        );

// Find the actual <input> inside the wrapper
        WebElement carrierInput = carrierInputWrapper.findElement(By.cssSelector("input"));

// Scroll into view
        ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", carrierInput);

// Optional wait
        Thread.sleep(500);

// Clear and enter the value
        carrierInput.clear();
        carrierInput.sendKeys(Carrier);

        System.out.println("Carrier ID " + Carrier + " entered using data-component-id.");


        //Visit Type

// Locate the ion-input wrapper using data-component-id
        WebElement visitTypeWrapper = driver.findElement(
                By.cssSelector("ion-input[data-component-id='VisitType']")
        );

// Find the inner <input> element inside the wrapper
        WebElement visitTypeInput = visitTypeWrapper.findElement(By.cssSelector("input"));

// Scroll into view to ensure it's clickable
        ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", visitTypeInput);

// Click the element
        visitTypeInput.click();

        System.out.println("VisitType field clicked.");
        Thread.sleep(1000);


        // Locate the button using its data-component-id
        WebElement liveUnloadButton = driver.findElement(
                By.cssSelector("button.dropdown-list[data-component-id='LiveLoad-dropdown-option']")
        );

// Scroll into view to ensure it's clickable
        ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", liveUnloadButton);

// Click the button
        liveUnloadButton.click();

        System.out.println("Clicked on 'Live Load' dropdown option.");


        Thread.sleep(1000);


        //Trailer

        // Locate the parent container first
        WebElement checkInSection = driver.findElement(By.cssSelector("div.checkInAttributes"));

// Then find the specific ion-input inside it using data-component-id
        WebElement trailerInputWrapper = checkInSection.findElement(
                By.cssSelector("ion-input[data-component-id='TrailerId']")
        );


// Find the actual <input> element inside the wrapper
        WebElement trailerInput = trailerInputWrapper.findElement(By.cssSelector("input"));


        // Scroll into view to ensure it's interactable
        ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", trailerInput);

// Optional: wait briefly to ensure stability
        Thread.sleep(500); // Replace with WebDriverWait if preferred

// Clear and enter the carrier value
        trailerInput.clear();
        trailerInput.sendKeys(Trailer);

        System.out.println("Trailer ID " + Trailer + " entered successfully.");


        Thread.sleep(1000);

        // Locate the ion-input wrapper using data-component-id
        WebElement equipmentTypeWrapper = driver.findElement(
                By.cssSelector("ion-input[data-component-id='EquipmentTypeId']")
        );

// Find the actual <input> element inside the wrapper
        WebElement equipmentTypeInput = equipmentTypeWrapper.findElement(By.cssSelector("input"));

// Scroll into view to ensure it's interactable
        ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", equipmentTypeInput);

// Optional: wait briefly to ensure stability
        Thread.sleep(500); // Replace with WebDriverWait if preferred

// Clear and enter the carrier value
        equipmentTypeInput.clear();
        equipmentTypeInput.sendKeys(Trailertype);

        System.out.println("Trailertype " + Trailertype + " entered successfully.");


        Thread.sleep(1000);

// Within that section, find the ion-input with data-component-id='LocationId'
        WebElement locationInputWrapper = checkInSection.findElement(
                By.cssSelector("ion-input[data-component-id='LocationId']")
        );

// Find the actual <input> element inside the wrapper
        WebElement locationInput = locationInputWrapper.findElement(By.cssSelector("input"));

// Scroll into view to ensure it's interactable
        ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", locationInput);

// Optional: wait briefly to ensure stability
        Thread.sleep(500); // Replace with WebDriverWait if preferred

// Clear and enter the carrier value
        locationInput.clear();
        locationInput.sendKeys(DockDoor);

        System.out.println("Location " + DockDoor + " entered successfully.");


        Thread.sleep(1000);
// Wait for the span element to be clickable
        WebDriverWait wait3 = new WebDriverWait(driver, Duration.ofSeconds(3));
        WebElement trailerContentDetails = wait3.until(ExpectedConditions.elementToBeClickable(
                By.cssSelector("span.expand-header-text[data-component-id='TrailerContentDetails']")
        ));

// Click the element
        trailerContentDetails.click();
        System.out.println("Trailer Content Details clicked.");


        js.executeScript("window.scrollTo(0, document.body.scrollHeight);");


        Thread.sleep(4000);


        // Within that section, find the ion-input with data-component-id='Shipment'
        WebElement shipmentInputWrapper = checkInSection.findElement(
                By.cssSelector("ion-input[data-component-id='Shipment']")
        );

// Find the actual <input> element inside the wrapper
        WebElement shipmentInput = shipmentInputWrapper.findElement(By.cssSelector("input"));

// Scroll into view in case it's off-screen or overlapped
        ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", shipmentInput);

// Add a short pause to ensure it's ready (optional but helpful)
        Thread.sleep(500); // or use WebDriverWait for better practice

// Clear and enter the shipment ID
        shipmentInput.clear();
        shipmentInput.sendKeys(shipmentNumber);

        System.out.println("Shipment ID " + shipmentNumber + " entered successfully.");


        // Locate the Check In button using its data-component-id
        WebElement checkInButton = driver.findElement(
                By.cssSelector("ion-button[data-component-id='checkin-btn']")
        );

// Scroll into view to ensure it's clickable
        ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", checkInButton);

// Click the button
        checkInButton.click();

        System.out.println("Clicked on 'Check In' button.");


    }
    public static void Load() throws InterruptedException {
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
        JavascriptExecutor js = (JavascriptExecutor) driver;
        Thread.sleep(10000);
        try {
            // 🔹 Step 1: Open the side menu
            WebElement shadowHost = driver.findElement(By.cssSelector("ion-button.menu-toggle-button"));
            SearchContext shadowRoot = (SearchContext) js.executeScript("return arguments[0].shadowRoot", shadowHost);
            WebElement menuButton = shadowRoot.findElement(By.cssSelector("button.button-native"));
            menuButton.click();

            // 🔹 Step 2: Search for "WM Mobile" in the menu
            WebElement searchInput = wait.until(ExpectedConditions.elementToBeClickable(
                    By.xpath("//input[@placeholder='Search Menu...']")));
            searchInput.clear();
            searchInput.sendKeys("WM Mobile");
//
            try {
                // Wait for the button to be present and visible
                WebElement wmMobileButton = wait.until(ExpectedConditions.presenceOfElementLocated(
                        By.cssSelector("button#wmMobile[data-component-id='WMMobile']")
                ));

                // Scroll into view and click using JS (in case normal click fails)
                js.executeScript("arguments[0].scrollIntoView(true);", wmMobileButton);
                js.executeScript("arguments[0].click();", wmMobileButton);
                System.out.println("✅ 'WM Mobile' button clicked.");
            } catch (Exception e) {
                System.err.println("❌ Error clicking 'WM Mobile': " + e.getMessage());
                e.printStackTrace();
            }
            // 🔹 Step 4: Switch to the new tab
            Thread.sleep(3000);
            ArrayList<String> tabs = new ArrayList<>(driver.getWindowHandles());
            driver.switchTo().window(tabs.get(1));
            Thread.sleep(8000);
            // 🔹 Step 5: Search for "Create Ilpn"


            WebElement searchInput1 = wait.until(ExpectedConditions.elementToBeClickable(
                    By.cssSelector("input.searchbar-input[placeholder='Search']")));
//                    By.xpath("//input[@type='search' and @placeholder='Search']")));
            searchInput1.click();
            searchInput1.clear();
            Thread.sleep(3000);
            searchInput1.sendKeys("JD Load Pallet-Static Route");
        } catch (Exception e) {
            System.err.println("❌ Error in Load: " + e.getMessage());
            e.printStackTrace();
        }
        // Locate the ion-label using its data-component-id
        WebElement labelElement = driver.findElement(
                By.cssSelector("ion-label[data-component-id='jdloadpallet-staticroute']")
        );

// Scroll into view to ensure it's interactable
        ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", labelElement);

// Click using JavaScript (in case native click doesn't work)
        ((JavascriptExecutor) driver).executeScript("arguments[0].click();", labelElement);

        System.out.println("Clicked on 'JD Load Pallet-Static Route' label.");


// Optional: wait briefly to ensure stability
        Thread.sleep(500); // Or use WebDriverWait if preferred


        WebElement dockDoorInput = new WebDriverWait(driver, Duration.ofSeconds(10))
                .until(ExpectedConditions.presenceOfElementLocated(
                        By.cssSelector("input[data-component-id*='dockdoorid']")
                ));


        WebElement dockDoorInput1 = new WebDriverWait(driver, Duration.ofSeconds(10))
                .until(ExpectedConditions.elementToBeClickable(
                        By.cssSelector("input[data-component-id='scandockdoor_togglefield_dockdoorid']")
                ));


        new WebDriverWait(driver, Duration.ofSeconds(10))
                .until(ExpectedConditions.visibilityOf(dockDoorInput1));


        dockDoorInput1.click();
        dockDoorInput1.clear();
        dockDoorInput1.sendKeys(DockDoor);


        System.out.println("DockDoor " + DockDoor + " entered successfully.");


        // Wait for the button to be clickable
        WebDriverWait wait1 = new WebDriverWait(driver, Duration.ofSeconds(10));
        WebElement goButton = wait1.until(ExpectedConditions.elementToBeClickable(
                By.cssSelector("ion-button[data-component-id='scandockdoor_togglefield_go']")
        ));

// Scroll into view if necessary
        ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", goButton);

// Click the button using JavaScript (useful for shadow DOM or complex components)
        ((JavascriptExecutor) driver).executeScript("arguments[0].click();", goButton);

        System.out.println("Go button clicked successfully.");


        // Wait for the input field to be visible and interactable
        WebDriverWait wait5 = new WebDriverWait(driver, Duration.ofSeconds(10));
        WebElement containerInput = wait5.until(ExpectedConditions.elementToBeClickable(
                By.cssSelector("input[data-component-id='scanolpn_barcodetextfield_olpn']")
        ));

// Scroll into view if needed
        ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", containerInput);

// Optional: click to activate the field
        containerInput.click();

// Clear and enter the value
        containerInput.clear();
        containerInput.sendKeys(pallet);

        System.out.println("Container value " + pallet + " entered successfully.");


        // Wait for the button to be clickable
        WebDriverWait wait6 = new WebDriverWait(driver, Duration.ofSeconds(10));

        WebElement goButton1 = wait6.until(ExpectedConditions.elementToBeClickable(
                By.cssSelector("ion-button[data-component-id='scanolpn_barcodetextfield_go']")
        ));

// Scroll into view if necessary
        ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", goButton1);

// Click the button using JavaScript (useful for shadow DOM or complex components)
        ((JavascriptExecutor) driver).executeScript("arguments[0].click();", goButton1);

        System.out.println("Go button clicked successfully.");


        Thread.sleep(5000); // Or use WebDriverWait if preferred


        // Wait until the ion-buttons container is present

        WebElement backButton = wait.until(ExpectedConditions.presenceOfElementLocated(
                By.cssSelector("ion-buttons[data-component-id='action_back_button']")
        ));

// Scroll into view to ensure visibility
        ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", backButton);

// Click the button twice using JavaScript

        js.executeScript("arguments[0].click();", backButton);
        Thread.sleep(300); // brief pause between clicks
        js.executeScript("arguments[0].click();", backButton);

        System.out.println("Back button clicked twice successfully.");


    }
    public static void JDCloseShipment() throws InterruptedException {
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
        JavascriptExecutor js = (JavascriptExecutor) driver;

//        // 🔹 Step 1: Open the side menu
//        WebElement shadowHost = driver.findElement(By.cssSelector("ion-button.menu-toggle-button"));
//        SearchContext shadowRoot = (SearchContext) js.executeScript("return arguments[0].shadowRoot", shadowHost);
//        WebElement menuButton = shadowRoot.findElement(By.cssSelector("button.button-native"));
//        menuButton.click();
//
//        // 🔹 Step 2: Search for "WM Mobile" in the menu
//        WebElement searchInput = wait.until(ExpectedConditions.elementToBeClickable(
//                By.xpath("//input[@placeholder='Search Menu...']")));
//        try {
//            searchInput.clear();
//            searchInput.sendKeys("WM Mobile");
//
//            try {
//                // Wait for the button to be present and visible
//                WebElement wmMobileButton = wait.until(ExpectedConditions.presenceOfElementLocated(
//                        By.cssSelector("button#wmMobile[data-component-id='WMMobile']")
//                ));
//
//                // Scroll into view and click using JS (in case normal click fails)
//                js.executeScript("arguments[0].scrollIntoView(true);", wmMobileButton);
//                js.executeScript("arguments[0].click();", wmMobileButton);
//                System.out.println("✅ 'WM Mobile' button clicked.");
//            } catch (Exception e) {
//                System.err.println("❌ Error clicking 'WM Mobile': " + e.getMessage());
//                e.printStackTrace();
//            }
//
//            // 🔹 Step 4: Switch to the new tab
//            Thread.sleep(3000);
//            ArrayList<String> tabs = new ArrayList<>(driver.getWindowHandles());
//            driver.switchTo().window(tabs.get(1));
//            Thread.sleep(5000);
//
//
//             //🔹 Step 5: Search for "Create Ilpn"
//

        try {
            WebElement searchInput1 = wait.until(ExpectedConditions.elementToBeClickable(
                    By.cssSelector("input.searchbar-input[placeholder='Search']")));
//                    By.xpath("//input[@type='search' and @placeholder='Search']")));
            searchInput1.click();
            searchInput1.clear();
            Thread.sleep(3000);
            searchInput1.sendKeys("JD Close Shipment");
        } catch (Exception e) {
            System.err.println("❌ Error in Close Shipment: " + e.getMessage());
            e.printStackTrace();
        }
        // Locate the ion-label using its data-component-id
        WebElement labelElement = driver.findElement(
                By.cssSelector("ion-label[data-component-id='jdcloseshipment']")
        );

// Scroll into view to ensure it's interactable
        ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", labelElement);

// Click using JavaScript (in case native click doesn't work)
        ((JavascriptExecutor) driver).executeScript("arguments[0].click();", labelElement);

        System.out.println("Clicked on 'jdcloseshipment");


// Optional: wait briefly to ensure stability
        Thread.sleep(500); // Or use WebDriverWait if preferred


        WebElement shipment1 = new WebDriverWait(driver, Duration.ofSeconds(10))
                .until(ExpectedConditions.elementToBeClickable(
                        By.cssSelector("input[data-component-id='scanshipment_barcodetextfield_shipment/bol']")
                ));


        new WebDriverWait(driver, Duration.ofSeconds(10))
                .until(ExpectedConditions.visibilityOf(shipment1));


        shipment1.click();
        shipment1.clear();
        shipment1.sendKeys(shipmentNumber);
        System.out.println("shipmentNumber " + shipmentNumber + " entered successfully.");


        // Wait for the button to be clickable
        WebDriverWait wait1 = new WebDriverWait(driver, Duration.ofSeconds(10));
        WebElement goButton = wait1.until(ExpectedConditions.elementToBeClickable(
                By.cssSelector("ion-button[data-component-id='scanshipment_barcodetextfield_go']")
        ));

// Scroll into view if necessary
        ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", goButton);

// Click the button using JavaScript (useful for shadow DOM or complex components)
        ((JavascriptExecutor) driver).executeScript("arguments[0].click();", goButton);

        System.out.println("Go button clicked successfully.");


        // Wait until the input field is visible and interactable

        WebElement sealInput = wait.until(ExpectedConditions.elementToBeClickable(
                By.cssSelector("input[data-component-id='custscanseal_barcodetextfield_shipmentsealnumber']")
        ));

// Scroll into view if needed
        ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", sealInput);

// Optional: click to activate the field
        sealInput.click();

// Clear and enter the value
        sealInput.clear();
        sealInput.sendKeys(sealinput);

        System.out.println("Shipment Seal value " + sealinput + " entered successfully.");


        // Wait until the button is clickable

        WebElement goButton3 = wait.until(ExpectedConditions.elementToBeClickable(
                By.cssSelector("ion-button[data-component-id='custscanseal_barcodetextfield_go']")
        ));

// Scroll into view if needed
        ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", goButton3);

// Click the button using JavaScript for reliability
        ((JavascriptExecutor) driver).executeScript("arguments[0].click();", goButton3);

        System.out.println("Go button clicked successfully.");


        Thread.sleep(1000);

        // Wait until the button is clickable

        WebElement endSealButton = wait.until(ExpectedConditions.elementToBeClickable(
                By.cssSelector("button[data-component-id='action_endshipmentseal_button']")
        ));

// Scroll into view to ensure visibility
        ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", endSealButton);

// Click the button using JavaScript for reliability
        ((JavascriptExecutor) driver).executeScript("arguments[0].click();", endSealButton);
        Thread.sleep(1000);
        ((JavascriptExecutor) driver).executeScript("arguments[0].click();", endSealButton);


        System.out.println("End Shipment Seal button clicked successfully.");


    }


}