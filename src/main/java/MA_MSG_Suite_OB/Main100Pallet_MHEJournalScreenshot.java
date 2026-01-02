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
import java.io.IOException;
import java.time.Duration;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import static MA_MSG_Suite_OB.DocPathManager.captureAll1;


public class Main100Pallet_MHEJournalScreenshot {
    public static WebDriver driver;
    public static int time = 60;
    public static class PalletReadyData {
        String PalletId;

        @Override
        public String toString() {
            return "PalletId: " + PalletId;
        }
    }
    public static class OLPNData {
        String olpn;
        @Override
        public String toString() {
            return "olpn: " + olpn ;
        }
    }

    public static void main(String Testcase, String filePath, String env, String messageType, String docPathLocal) throws IOException {
  //  public static void main(String[] args) {

        List<PalletReadyData> palletList = readPalletReadyData(filePath, Testcase);

        System.out.println("Testcase:"+Testcase);
        System.out.println("filePath:"+filePath);

        try {
            WebDriverManager.chromedriver().setup();
            ChromeOptions options = new ChromeOptions();
            options.addArguments("--start-maximized");
            driver = new ChromeDriver(options);
            driver.manage().window().maximize();
            Main1_URL_Login1 login1 = new Main1_URL_Login1(driver, env);
            login1.execute();
            System.out.println("login done:\n");

            SearchMenu("MHE Journal","whseDeviceIntegrationMessageJournal");

            System.out.println("OpenFilter 1");
            Thread.sleep(5000);
            OpenFilter();
            System.out.println("ChevronDown 1");
            ChevronDown("MessageJournal-MessageType-chevron-down");
            System.out.println("EnterValue 1");
            EnterValue(messageType,"MessageType");





//                List<OLPNData> payLoadList = readOlpnData1(filePath, Testcase);

                System.out.println("\n✅ Starting OLPN loop...");
                for (PalletReadyData data : palletList) {

                    String PalletId = data.PalletId;
                    System.out.println("\n🔍 Searching for OLPN: " + PalletId);
                    Thread.sleep(3000);

                    // Open filter
                    System.out.println("OpenFilter 2");
                    OpenFilter();
                    Thread.sleep(5000);

                    // Expand payload dropdown
                    System.out.println("ChevronDown 2");
                    ChevronDown("MessageJournal-Payload-chevron-down");
                    Thread.sleep(10000);

                    // Enter OLPN value
                    System.out.println("EnterValue 2");
                    EnterValue(PalletId, "Stage1.MessagePayload");
                    Thread.sleep(2000);

                    // Select first row
                    tickFirstRowCheckbox();
                    Thread.sleep(1000);

                    //Capture Screenshot
                    DocPathManager.captureScreenshot("Message ", driver);
                    DocPathManager.saveSharedDocument();
                    Thread.sleep(3000);

                    // Open details
                    clickDetailsButton();
                    Thread.sleep(2000);

                    // Open details icon
                    clickFirstRowDetailsIcon(docPathLocal);
                    Thread.sleep(5000);

                    System.out.println("✅ Completed processing for OLPN: " + PalletId);

                    //Addd screenshot method here
                    captureAll1(filePath, driver);
                    DocPathManager.saveSharedDocument();


                    findCloseButtonInModal();
                    Thread.sleep(2000);

                    navigateTillMHEJournal();

                }





                System.out.println("Executing flow for messageType: " + messageType);



              //  Main100_OlpnScreenShot.main(filePath, Testcase, driver, env, docPathLocal);









        } catch (Exception e) {
            System.err.println(" Error: " + e.getMessage());
            e.printStackTrace();
        }finally {
            System.out.println("MHE Journal Done");
            if (driver != null) driver.quit();

            if (driver == null) {
                System.out.println("Driver is NULL11");
            } else {
                System.out.println("Driver is initialized");
            }


        }
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

//
//        List<WebElement> closeButtons = driver.findElements(By.id("close-menu-button"));
//
//        if (!closeButtons.isEmpty() && closeButtons.get(0).isDisplayed()) {
//            System.out.println("GOT");
//
//            // If present and visible, click it
//            closeButtons.get(0).click();
//        }




    }

    public static void OpenFilter() throws InterruptedException {
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(time));

        try {
            WebElement filterBtnHost = wait.until(ExpectedConditions.elementToBeClickable(
                    By.xpath("(//ion-button[contains(@class,\"toggle-button\")])[3]")
            ));
            filterBtnHost.click();
            System.out.println("Click on Filter button in MHE Journal ");
        } catch (Exception e) {
            System.err.println("Error: " + e.getMessage());
            e.printStackTrace(System.err);
        }


    }

    public static void ChevronDown(String datacomponentid) throws InterruptedException {
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(time));
        JavascriptExecutor js = (JavascriptExecutor) driver;

        //  chevron-down (component-ids can vary)
        try {
            WebElement expandButton1 = wait.until(
                    ExpectedConditions.elementToBeClickable(
                            By.cssSelector("ion-button[data-component-id='"+datacomponentid+"']")
                    )
            );
            ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", expandButton1);

            // System.out.println("Wave Run chevron-down button clicked in  sec");

            Thread.sleep(3000);
            expandButton1.click();
            System.out.println("MHE Journal chevron-down button clicked using native click.");

        }
        catch (Exception e) {
            System.err.println("Error: " + e.getMessage());
            e.printStackTrace(System.err);
        }



    }

    public static void EnterValue(String message, String ID) throws InterruptedException {
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(time));
        JavascriptExecutor js = (JavascriptExecutor) driver;
        try {
            WebElement input = driver.findElement(
                    By.cssSelector("ion-input[data-component-id='"+ID+"'] input")
            );
            input.click();
            input.clear();
            input.sendKeys(message);
            input.sendKeys(Keys.ENTER);
            System.out.println("Message Type entered in Filter button.");
        } catch (Exception e) {
            System.err.println("Error: " + e.getMessage());
            e.printStackTrace(System.err);
        }



    }

    public static void freeEnterValue( String ID) throws InterruptedException {
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(time));
        JavascriptExecutor js = (JavascriptExecutor) driver;
        try {
            WebElement input = driver.findElement(
                    By.cssSelector("ion-input[data-component-id='"+ID+"'] input")
            );
            input.click();
            input.sendKeys(Keys.ENTER);
            System.out.println("Message Type entered in Filter button.");
        } catch (Exception e) {
            System.err.println("Error: " + e.getMessage());
            e.printStackTrace(System.err);
        }



    }

    public static void tickFirstRowCheckbox() {
        int maxRetries = 10;   // safety limit to avoid infinite loop
        int attempts = 0;

        try {
            while (attempts < maxRetries) {

                // 1. Locate all rows
                List<WebElement> rows = driver.findElements(
                        By.cssSelector("datatable-body-row.datatable-body-row")
                );

                // 2. If no rows OR first row not visible → open filter and retry
                if (rows.isEmpty() || !rows.get(0).isDisplayed()) {
                    System.out.println("⚠️ First row not visible. Opening filter… Attempt: " + (attempts + 1));
                    OpenFilter();
                    Thread.sleep(1000);

                    System.out.println("Checking Row again");
                    Thread.sleep(5000);

                    freeEnterValue("Stage1.MessagePayload");
                    System.out.println("Wait for 10 sec");

                    Thread.sleep(10000); // small wait for UI refresh
                    attempts++;
                    continue;
                }

                // 3. First row is visible → click checkbox
                WebElement firstRow = rows.get(0);

                WebElement checkbox = firstRow.findElement(
                        By.cssSelector("label.datatable-checkbox input[type='checkbox']")
                );

                ((JavascriptExecutor) driver).executeScript("arguments[0].click();", checkbox);

                System.out.println("✅ First row checkbox selected");
                return;
            }

            System.out.println("❌ Failed to find a visible first row after " + maxRetries + " attempts.");

        } catch (Exception e) {
            System.out.println("⚠️ Error selecting first row checkbox: " + e.getMessage());
        }
    }


    public static void clickDetailsButton() {
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(time));
        try {
            // 1. Wait until the ion-button is visible in DOM
            WebElement detailsBtn = wait.until(
                    ExpectedConditions.visibilityOfElementLocated(
                            By.cssSelector("ion-button[data-component-id='footer-panel-action-Details']")
                    )
            );

            // 2. Click using JS (Ionic Shadow DOM safe)
            ((JavascriptExecutor) driver).executeScript("arguments[0].click();", detailsBtn);

            System.out.println("✅ DETAILS button clicked");

        } catch (Exception e) {
            System.out.println("⚠️ Unable to click DETAILS button: " + e.getMessage());
        }
    }

    public static void clickFirstRowDetailsIcon(String docPathLocal) throws InterruptedException {
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(time));
        try {
            // Wait until the element is visible
            WebElement icon = wait.until(
                    ExpectedConditions.visibilityOfElementLocated(
                            By.cssSelector("div[data-component-id='screen-header-card-inactive']")
                    )
            );

            // Click using JavaScript (safer for Angular/Ionic)
            ((JavascriptExecutor) driver).executeScript("arguments[0].click();", icon);

            System.out.println("✅ Header card inactive icon clicked");

        } catch (Exception e) {
            System.out.println("⚠️ Unable to click header card inactive icon: " + e.getMessage());
        }
        try {
            // 1. Wait for all cards to be present
            List<WebElement> cards = wait.until(
                    ExpectedConditions.presenceOfAllElementsLocatedBy(
                            By.cssSelector("card-view[data-component-id='Card-View']")
                    )
            );

            if (cards.isEmpty()) {
                System.out.println("⚠️ No cards found in panel");
                return;
            }

            // 2. Get the first card
            WebElement firstCard = cards.get(0);

            // 3. Locate the details button inside the first card
            WebElement detailsButton = firstCard.findElement(
                    By.cssSelector("button.details-button[data-component-id='details-button']")
            );

            // 4. Wait until visible
            wait.until(ExpectedConditions.visibilityOf(detailsButton));

            // 5. Click using JS (SVG + Shadow DOM safe)
            ((JavascriptExecutor) driver).executeScript("arguments[0].click();", detailsButton);

            System.out.println("✅ Details icon clicked for first card");

        } catch (Exception e) {
            System.out.println("⚠️ Error clicking details icon: " + e.getMessage());
        }

        Thread.sleep(3000);
        System.out.println("Path: "+ docPathLocal);
        DocPathManager.captureScreenshot("Message Payload ",driver);
        DocPathManager.saveSharedDocument();
        System.out.println("Message Payload  screenshot done");
        Thread.sleep(3000);



    }

    public static void findCloseButtonInModal() {
        WebElement closeBtn = driver.findElement(
                By.xpath("//ion-button[normalize-space()='Close']")
        );
        closeBtn.click();
    }

    public static void navigateTillMHEJournal() throws InterruptedException, IOException {
        WebElement waveRunsLink = driver.findElement(By.cssSelector("a[title='MHE Journal']"));
        ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", waveRunsLink);
        waveRunsLink.click();
        Thread.sleep(5000);
        System.out.println("MHE Journal clicked.");
    }

    public static List<OLPNData> readOlpnData1(String path, String testcase) throws IOException {
        List<OLPNData> list = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(path);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheet("Tasks");
            if (sheet == null) {
                System.err.println("❌ Sheet 'Tasks' not found.");
                return list;
            }
            System.out.println("Testcase No."+ testcase);
            DataFormatter fmt = new DataFormatter();

            for (int i = 1; i <= sheet.getLastRowNum(); i++) { // assume row 0 is header
                Row row = sheet.getRow(i);
                if (row == null) continue;

                String testcaseCell = fmt.formatCellValue(row.getCell(0)).trim();   // Testcase @ col 1
                if (!testcaseCell.equalsIgnoreCase(testcase == null ? "" : testcase.trim())) {
                    // System.err.println("No Testcase Found");
                    continue; // only take rows matching the input Testcase (e.g., "TST_001")

                }

                String olpn = fmt.formatCellValue(row.getCell(4)).trim();     // WCSOrderId @ col 4

                if (!olpn.isEmpty() ) {
                    OLPNData data = new OLPNData();
                    data.olpn = olpn;
                    list.add(data);
                } else {
                    throw new IllegalArgumentException("Invalid data: please check LCID and WCS Order ID");
                }

            }
        }

        return list;
    }


//    public static List<PalletReadyData> readPalletReadyData(String path, String testcase) throws IOException {
//        List<PalletReadyData> list = new ArrayList<>();
//
//        try (FileInputStream fis = new FileInputStream(path);
//             Workbook workbook = new XSSFWorkbook(fis)) {
//
//            Sheet sheet = workbook.getSheet("Tasks");
//            if (sheet == null) {
//                System.err.println("❌ Sheet 'Tasks' not found.");
//                return list;
//            }
//
//            DataFormatter fmt = new DataFormatter();
//
//            // ✅ Find the column index for "Pallet"
//            int palletCol = getColumnIndexByName(sheet, "Pallet");
//            if (palletCol == -1) {
//                throw new RuntimeException("❌ Column 'Pallet' not found in Tasks sheet header.");
//            }
//
//            // Testcase filter logic
//            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
//                Row row = sheet.getRow(i);
//                if (row == null) continue;
//
//                String testcaseCell = fmt.formatCellValue(row.getCell(0)).trim();
//                if (!testcaseCell.equalsIgnoreCase(testcase.trim())) continue;
//
//                // ✅ Read using column name
//                String palletId = fmt.formatCellValue(row.getCell(palletCol)).trim();
//
//                if (!palletId.isEmpty()) {
//                    PalletReadyData data = new PalletReadyData();
//                    data.PalletId = palletId;
//                    list.add(data);
//                }
//            }
//        }
//        return list;
//    }
//




    private static List<PalletReadyData> readPalletReadyData(String path, String testcase) throws IOException {
        List<PalletReadyData> list = new ArrayList<>();
        Set<String> uniquePallets = new HashSet<>(); // ✅ Track uniqueness

        try (FileInputStream fis = new FileInputStream(path);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheet("Tasks");
            if (sheet == null) {
                System.err.println("❌ Sheet 'Tasks' not found.");
                return list;
            }

            DataFormatter fmt = new DataFormatter();

            // ✅ Find the column index for "Pallet"
            int palletCol = getColumnIndexByName(sheet, "Pallet");
            if (palletCol == -1) {
                throw new RuntimeException("❌ Column 'Pallet' not found in Tasks sheet header.");
            }

            // Testcase filter logic
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                String testcaseCell = fmt.formatCellValue(row.getCell(0)).trim();
                if (!testcaseCell.equalsIgnoreCase(testcase.trim())) continue;

                // Read using column name
                String palletId = fmt.formatCellValue(row.getCell(palletCol)).trim();

                if (!palletId.isEmpty() && uniquePallets.add(palletId)) {
                    // add() returns false if palletId already exists
                    PalletReadyData data = new PalletReadyData();
                    data.PalletId = palletId;
                    list.add(data);
                }
            }
        }
        return list;
    }



    private static int getColumnIndexByName(Sheet sheet, String columnName) {
        Row headerRow = sheet.getRow(0);  // header is in row 0
        if (headerRow == null) return -1;

        DataFormatter formatter = new DataFormatter();

        for (int i = 0; i < headerRow.getLastCellNum(); i++) {
            Cell cell = headerRow.getCell(i);
            if (cell == null) continue;

            String header = formatter.formatCellValue(cell).trim();
            if (header.equalsIgnoreCase(columnName.trim())) {
                return i;   // found!
            }
        }
        return -1; // not found
    }






}








