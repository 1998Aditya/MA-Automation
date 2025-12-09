package MA_MSG_Suite_OB;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.openqa.selenium.support.ui.ExpectedConditions;
import io.github.bonigarcia.wdm.WebDriverManager;

import java.io.FileInputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.ArrayList;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Main11_Shipment {

    public static WebDriver driver;

    // Excel-driven variables
    public static String pallet;
    public static String staticRoute;
    public static String location;
    public static String carrier;
    public static String trailer;
    public static String trailerType;
    public static String dockDoor;
    public static String sealInput;
    public static String shipmentNumber;
   // public static String olpn = "KU700311001557";

    public static void main( String filePath, String sheetName, String testcaseId, String env) throws InterruptedException {
//        String filePath = "C:\\Users\\2210420\\IdeaProjects\\msg-runner\\OOdata.xlsx";
//        String sheetName = "Outbound";
//        String testcaseId = "TST_001"; // 🔹 The testcase you want to run

        runAllMatchingTestcases(filePath, sheetName, testcaseId,env);
    }

    public static void runAllMatchingTestcases(String filePath, String sheetName, String testcaseId, String  env) {
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheet(sheetName);
            int matchCount = 0;

            for (Row row : sheet) {
                Cell testcaseCell = row.getCell(0); // Column 0 = testcase ID
                System.out.println("Row " + row.getRowNum() + " Col0 value: " + row.getCell(0));

                if (testcaseCell != null) {
                    String cellValue = testcaseCell.toString().trim();
                    System.out.println("Checking row " + row.getRowNum() + ": " + cellValue);

                    if (cellValue.equalsIgnoreCase(testcaseId.trim())) {
                        matchCount++;
                        System.out.println("\n🔁 Running " + testcaseId + " for row " + row.getRowNum());

                        // Load variables from this row
                        pallet       = getCellValue(row.getCell(1));
                        staticRoute  = getCellValue(row.getCell(2));
                        location     = getCellValue(row.getCell(3));
                        carrier      = getCellValue(row.getCell(4));
                        trailer      = getCellValue(row.getCell(5));
                        trailerType  = getCellValue(row.getCell(6));
                        dockDoor     = getCellValue(row.getCell(7));
                        sealInput    = getCellValue(row.getCell(8));

                        printVariables();

                        // Run your flow for this row

                        WebDriverManager.chromedriver().setup();
                        ChromeOptions options = new ChromeOptions();
                        options.addArguments("--start-maximized");

                        driver = new ChromeDriver(options);
                        driver.manage().window().maximize();
                        Main1_URL_Login1 login1 = new Main1_URL_Login1(driver, env);
                        login1.execute();
                        System.out.println("login done:\n");


                        Shipment();
        CheckIn();
       Load();
       JDCloseShipment();

                        driver.quit(); // Close browser after each run
                    }
                }
            }

            if (matchCount == 0) {
                System.out.println("❌ No matching rows found for testcase: " + testcaseId);
            }

        } catch (Exception e) {
            System.err.println("❌ Error processing Excel: " + e.getMessage());
            e.printStackTrace();
        }
    }

    private static String getCellValue(Cell cell) {
        return cell == null ? "" : cell.toString().trim();
    }

    public static void printVariables() {
        System.out.println("🔍 Loaded Testcase Variables:");
        System.out.println("Pallet: " + pallet);
        System.out.println("StaticRoute: " + staticRoute);
        System.out.println("Location: " + location);
        System.out.println("Carrier: " + carrier);
        System.out.println("Trailer: " + trailer);
        System.out.println("TrailerType: " + trailerType);
        System.out.println("DockDoor: " + dockDoor);
        System.out.println("SealInput: " + sealInput);
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
        if (staticRoute != null && !staticRoute.isEmpty()) {
            orderPlanningRunInputField.click();
            orderPlanningRunInputField.clear();
            orderPlanningRunInputField.sendKeys(staticRoute);
            orderPlanningRunInputField.sendKeys(Keys.ENTER);
        }

        System.out.println("StaticRoute entered: " + staticRoute);

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
        WebDriverWait wait2 = new WebDriverWait(driver, Duration.ofSeconds(10));
        WebElement checkInInfo = wait2.until(ExpectedConditions.elementToBeClickable(
                By.cssSelector("span.expand-header-text[data-component-id='Check-InInfo']")
        ));

// Click the element
        checkInInfo.click();
        System.out.println("Check-In Info clicked.");


        // Locate the ion-input wrapper using data-component-id
        WebElement carrierInputWrapper = driver.findElement(
                By.cssSelector("ion-input[data-component-id='carrierId']")
        );

// Find the actual <input> inside the wrapper
        WebElement carrierInput = carrierInputWrapper.findElement(By.cssSelector("input"));

// Scroll into view
        ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", carrierInput);

// Optional wait
        Thread.sleep(500);

// Clear and enter the value
        carrierInput.clear();
        carrierInput.sendKeys(carrier);

        System.out.println("carrier ID " + carrier + " entered using data-component-id.");


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
        trailerInput.sendKeys(trailer);

        System.out.println("Trailer ID " + trailer + " entered successfully.");


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
        equipmentTypeInput.sendKeys(trailerType);

        System.out.println("Trailertype " + trailerType + " entered successfully.");


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
        locationInput.sendKeys(dockDoor);

        System.out.println("Location " + dockDoor + " entered successfully.");


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
        dockDoorInput1.sendKeys(dockDoor);


        System.out.println("DockDoor " + dockDoor + " entered successfully.");


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
        sealInput.sendKeys((CharSequence) sealInput);

        System.out.println("Shipment Seal value " + sealInput + " entered successfully.");


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