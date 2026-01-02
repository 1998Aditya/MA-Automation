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
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;


public class Main8_B2BPackPickCart {
    public static int time =60;
    public static WebDriver driver;
    public static String docPathLocal ;
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
        // SearchInWMMobile("JD Pack Pick Cart");//", "jdpackpickcart");
        Thread.sleep(5000);
        SearchInWmMobile("JD Pack Pick Cart","jdpackpickcart");
        //Whatever present in Column E it will pack
       // XWPFDocument document = new XWPFDocument();
         docPathLocal = DocPathManager.getOrCreateDocPath(filePath, testcase);

        Packing(filePath,testcase);

        SwitchTab.tabswitch(driver, 0);

        System.out.println("Wait 5 sec");
        Thread.sleep(5000);
        Main100_OlpnScreenShot.main(filePath, testcase,driver, env,docPathLocal);
        SwitchTab.tabswitch(driver, 1);
        Thread.sleep(5000);



        if (driver != null) {
        driver.quit();
    }
        System.out.println(" B2BPackPickCart is Done for Testcase"+testcase);

}


    public static void SearchMenuWM(String Keyword, String id)  {
        WebDriverWait wait1 = new WebDriverWait(driver, Duration.ofSeconds(time));
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(time));
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
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(time));


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
    public static WebElement getIonSearchBarInput() {
        JavascriptExecutor js = (JavascriptExecutor) driver;
        return (WebElement)
                js.executeScript("return document.querySelector('ion-searchbar')"
                        + ".shadowRoot.querySelector('input.searchbar-input');"
                );
    }



    public static void Packing(String filePath, String testcaseName) {
        System.out.println("✅ Search Done: " + filePath);
        System.out.println("✅ Search Done: " + testcaseName);

        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(time));
        JavascriptExecutor js = (JavascriptExecutor) driver;

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(3); // sheet index 3

            List<String> olpns = new ArrayList<>();

            // Iterate rows (skip header row)
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                String testcase = getCellString(row.getCell(0)); // Column A
                if (testcase == null || testcase.isEmpty()) continue; // skip blanks

                if (testcase.equalsIgnoreCase(testcaseName)) {

                    String olpn = getCellString(row.getCell(4)); // Column E
                    if (olpn != null && !olpn.isEmpty()) {
                        olpns.add(olpn.trim());
                    }
                }
            }

            // No need to call workbook.close(); try-with-resources handles it

            if (olpns.isEmpty()) {
                System.out.println("❌ No OLPNs found for testcase: " + testcaseName);
                return;
            }

            System.out.println("=== Processing Testcase: " + testcaseName + " ===");

            for (String olpn : olpns) {
                System.out.println("Packing OLPN: " + olpn);

                WebElement skipPrinterScanButton = wait.until(
                        ExpectedConditions.elementToBeClickable(
                                By.cssSelector("button[data-component-id='action_skipprinterscan_button']")
                        )
                );
                skipPrinterScanButton.click();

                WebElement olpnInput = wait.until(
                        ExpectedConditions.visibilityOfElementLocated(
                                By.cssSelector("input[data-component-id='acceptolpn_barcodetextfield_olpn']")
                        )
                );
                olpnInput.clear();
                olpnInput.sendKeys(olpn);
                Thread.sleep(3000);
                DocPathManager.captureScreenshot("Olpn ",driver);
                DocPathManager.saveSharedDocument();
                System.out.println("Olpn screenshot done");
                Thread.sleep(3000);
                olpnInput.sendKeys(Keys.ENTER);

                try {
                    WebDriverWait shortWait = new WebDriverWait(driver, Duration.ofSeconds(time));
                    WebElement okButton = shortWait.until(
                            ExpectedConditions.elementToBeClickable(By.xpath("//button[.//span[text()='Ok']]"))
                    );
                    js.executeScript("arguments[0].scrollIntoView(true);", okButton);
                    Thread.sleep(3000);
                    DocPathManager.captureScreenshot("Click OK ",driver);
                    DocPathManager.saveSharedDocument();
                    System.out.println("Click OK screenshot done");
                    Thread.sleep(3000);
                    js.executeScript("arguments[0].click();", okButton);
                    System.out.println("Clicked on 'Ok' button in alert.");
                } catch (ElementClickInterceptedException ex) {
                    System.out.println("Alert not found or click failed. Please click manually.");
                    Thread.sleep(10000);
                    System.out.println("10 sec remaining");
                    Thread.sleep(10000);
                }
            }



            WebElement backButton = wait.until(
                    ExpectedConditions.presenceOfElementLocated(
                            By.cssSelector("ion-buttons[data-component-id='action_back_button']")
                    )
            );
            js.executeScript("arguments[0].scrollIntoView(true);", backButton);
            Thread.sleep(300);
            js.executeScript("arguments[0].click();", backButton);

            System.out.println("Back button clicked successfully.");
            Thread.sleep(3000);

        } catch (Exception e) {
            System.out.println("❌ Exception in Packing(): " + e.getMessage());
            e.printStackTrace();
        }
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




//public static void captureScreenshot(String fileName) {
//    try {
//        File srcFile = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
//        try (FileInputStream fis = new FileInputStream(srcFile)) {
//            XWPFDocument document = DocPathManager.getSharedDocument();
//            XWPFParagraph paragraph = document.createParagraph();
//            XWPFRun run = paragraph.createRun();
//            run.setText("Screenshot: " + fileName);
//            run.addBreak();
//            run.addPicture(fis,
//                    Document.PICTURE_TYPE_PNG,
//                    fileName + ".png",
//                    Units.toEMU(500),
//                    Units.toEMU(300));
//        }
//        System.out.println("Screenshot added to document: " + fileName);
//    } catch (Exception e) {
//        System.out.println("Error capturing screenshot: " + e.getMessage());
//    }
//}
//    public static void captureAllCardsScreenshots() throws InterruptedException, IOException {
//        XWPFDocument document = DocPathManager.getSharedDocument(); // shared doc
//        List<WebElement> rows = driver.findElements(By.cssSelector("[role='main'] card-view"));
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



}