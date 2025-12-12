package MA_MSG_Suite_OB;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.FileInputStream;
import java.time.Duration;
import java.util.ArrayList;


public class Main9_Palletisation {
    public static String location;// = "CONHUBALICANTE";
    public static String pallet;// = "CONHUBALICANTE";
    public static WebDriver driver;
    public static void main(String filePath, String testcase, String env) throws InterruptedException {
        WebDriverManager.chromedriver().setup();
        ChromeOptions options = new ChromeOptions();
        options.addArguments("--start-maximized");

        driver = new ChromeDriver(options);
        driver.manage().window().maximize();
        Main1_URL_Login1 login1 = new Main1_URL_Login1(driver, env);
        login1.execute();
        System.out.println("login done:\n");
        SearchMenuWM("WM Mobile","WMMobile");
        //SearchInWmMobile("JD OB Putaway To Staging", "jdobputawaytostaging");
        //Whatever present in Column E it will pack
       // OutboundPutaway(filePath,testcase);

    }

    public static void SearchMenuWM(String Keyword, String id)  {
        WebDriverWait wait1 = new WebDriverWait(driver, Duration.ofSeconds(15));
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
        JavascriptExecutor js = (JavascriptExecutor) driver;




        int maxRetries = 6; // Try up to 2 times (1 initial + 1 retry)
        for (int attempt = 1; attempt <= maxRetries; attempt++) {
            System.out.println("⏳ Waiting 10 seconds before refreshing...");

            try {
                Thread.sleep(10000); // Wait 1 minute
            } catch (InterruptedException e) {
                throw new RuntimeException(e);
            }

//                // 🔄 Click the refresh button inside shadow DOM
//                WebElement refreshHost = driver.findElement(By.cssSelector("ion-button.refresh-button"));
//                js = (JavascriptExecutor) driver;
//                WebElement refreshButton = (WebElement) js.executeScript(
//                        "return arguments[0].shadowRoot.querySelector('button.button-native')", refreshHost);
//                refreshButton.click();
//                System.out.println("🔄 Refresh button clicked.");

            // Locate using data-component-id
            WebElement refreshBtn = wait.until(
                    ExpectedConditions.elementToBeClickable(
                            By.cssSelector("ion-button[data-component-id='refresh']")
                    )
            );

            // Click the button
            refreshBtn.click();

            // Optional: verify action or add logging
            System.out.println("Refresh button clicked successfully!");

            // Optional: wait for UI to settle
            try {
                Thread.sleep(3000);
            } catch (InterruptedException e) {
                throw new RuntimeException(e);
            }

            try {
                WebElement shadowHost = wait1.until(ExpectedConditions.presenceOfElementLocated(
                        By.cssSelector("ion-button[data-component-id='menu-toggle-button']")
                ));
                SearchContext shadowRoot = (SearchContext) js.executeScript("return arguments[0].shadowRoot", shadowHost);
                WebElement nativeButton = shadowRoot.findElement(By.cssSelector("button.button-native"));

// wait for overlay to disappear
                wait1.until(ExpectedConditions.invisibilityOfElementLocated(By.cssSelector("manh-overlay-container")));

// click via JS to avoid interception
                js.executeScript("arguments[0].click();", nativeButton);

                System.out.println("Menu toggle button clicked.");




                break;
            }catch (Exception e) {
                System.err.println("Error: " + e.getMessage());
                e.printStackTrace(System.err);


            }


        }




        try {
            // Locate the inner input directly under ion-input
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


            // Wait for the button to be present and visible
            WebElement element = wait.until(ExpectedConditions.presenceOfElementLocated(
                    By.cssSelector("button#wmMobile[data-component-id=" + id + "]")
            ));


            ((JavascriptExecutor) driver).executeScript(
                    "arguments[0].scrollIntoView({block: 'center'});", element
            );

            ((JavascriptExecutor) driver).executeScript(
                    "arguments[0].click();", element
            );

            System.out.println("Clicked element with id: " + id);
            ArrayList<String> tabs = new ArrayList<>(driver.getWindowHandles());
            driver.switchTo().window(tabs.get(1));

        } catch (Exception e) {

            System.err.println("Failed to click element with id: " + id);
            e.printStackTrace();

        }


    }


    static void SearchInWmMobile(String Transaction, String ComponentId)  {
        JavascriptExecutor js = (JavascriptExecutor) driver;
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));


        try {


            WebElement searchInput = wait.until(ExpectedConditions.elementToBeClickable(
                    By.cssSelector("ion-searchbar[data-component-id='search'] input[type='search']")));

            // Clear any existing text
            searchInput.clear();

            // Type the search text
            searchInput.sendKeys(Transaction);

            // Optionally, press ENTER if the search requires submission
            searchInput.sendKeys(Keys.ENTER);


//            WebElement searchInput1 = wait.until(ExpectedConditions.elementToBeClickable(
//                    By.cssSelector("input.searchbar-input[placeholder='Search']")));
////                    By.xpath("//input[@type='search' and @placeholder='Search']")));
//            searchInput1.click();
//            searchInput1.clear();
//            Thread.sleep(3000);
//            searchInput1.sendKeys(Transaction);
        } catch (Exception e) {
            System.err.println("❌ Error in "+Transaction + e.getMessage());
            e.printStackTrace();
        }
        // Locate the ion-label using its data-component-id

        WebElement labelElement = wait.until(ExpectedConditions.elementToBeClickable(
                By.cssSelector("ion-label[data-component-id='" + ComponentId + "']")
        ));

        // Scroll into view to ensure it's interactable
        ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", labelElement);

        // Click using JavaScript (in case native click doesn't work)
        ((JavascriptExecutor) driver).executeScript("arguments[0].click();", labelElement);

        System.out.println("Clicked on" +Transaction+" label.");


    }

    public static void OutboundPutaway(String filePath,String testcase) throws InterruptedException {
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
        JavascriptExecutor js = (JavascriptExecutor) driver;

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheet("Outbound");

            // Get Row 1 (index 0 because POI is zero-based)
            Row row = sheet.getRow(0);

            // Map cells to variables
            String pallet = row.getCell(0).getStringCellValue();  // Row 1, Col 1
            String StaticRoute = row.getCell(1).getStringCellValue();  // Row 1, Col 2
            String location = row.getCell(2).getStringCellValue();  // Row 1, Col 3
            //   int quantity = (int) row.getCell(3).getNumericCellValue(); // Row 1, Col 4

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

    /**
     * Safely get a trimmed String from a cell.
     * Returns null for null/blank cells.
     */
    private static String getCellString(Cell cell) {
        if (cell == null) return null;

        switch (cell.getCellType()) {
            case STRING:
                String s = cell.getStringCellValue();
                return (s == null) ? null : s.trim();
            case NUMERIC:
                // Convert numeric to string without scientific format
                return new java.text.DecimalFormat("#.################").format(cell.getNumericCellValue()).trim();
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue()).trim();
            case FORMULA:
                // Evaluate cached value type
                CellType cached = cell.getCachedFormulaResultType();
                if (cached == CellType.STRING) {
                    String fs = cell.getStringCellValue();
                    return (fs == null) ? null : fs.trim();
                } else if (cached == CellType.NUMERIC) {
                    return new java.text.DecimalFormat("#.################").format(cell.getNumericCellValue()).trim();
                } else if (cached == CellType.BOOLEAN) {
                    return String.valueOf(cell.getBooleanCellValue()).trim();
                } else {
                    return null;
                }
            case BLANK:
            case _NONE:
            case ERROR:
            default:
                return null;
        }
    }

//// Scroll into view to ensure visibility
//            ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", backButton);
//
//// Click the button twice using JavaScript
//
//            // js.executeScript("arguments[0].click();", backButton);
//            Thread.sleep(300); // brief pause between clicks
//            js.executeScript("arguments[0].click();", backButton);
//
//            System.out.println("Back button clicked twice successfully.");
//            Thread.sleep(3000);
//        } catch (Exception e) {
//            System.out.println("❌ Exception in wmmobile(): " + e.getMessage());
//            e.printStackTrace();
//        }
//
//    }
}