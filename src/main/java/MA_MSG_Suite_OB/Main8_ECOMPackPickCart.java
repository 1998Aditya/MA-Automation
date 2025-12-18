package MA_MSG_Suite_OB;


import okhttp3.*;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.json.JSONArray;
import org.json.JSONObject;
import org.openqa.selenium.*;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.File;
import java.io.FileOutputStream;
import java.util.*;

import com.google.gson.JsonObject;
import com.google.gson.JsonParser;
import io.github.bonigarcia.wdm.WebDriverManager;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

import java.io.FileInputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

import static MA_MSG_Suite_OB.Main7_MakePickCart.searchItemMaster;


public class Main8_ECOMPackPickCart {
//    public static String filePath = "C:\\Users\\2210420\\IdeaProjects\\msg-runner\\OOdata.xlsx";
//    public static String testcase = "TST_001"; // You can pass this dynamically
//    public static String env ="DEV";


    public static int time =60;
    public static WebDriver driver;
    public static String docPathLocal ;

    //  public static void main(String[] args) throws InterruptedException {
    public static void main( String filePath,String testcase, String env) throws InterruptedException{


        WebDriverManager.chromedriver().setup();
        ChromeOptions options = new ChromeOptions();
        options.addArguments("--start-maximized");

        driver = new ChromeDriver(options);
        driver.manage().window().maximize();
        Main1_URL_Login1 login1 = new Main1_URL_Login1(driver, env);
        login1.execute();
        System.out.println("login done:\n");
        docPathLocal = DocPathManager.getOrCreateDocPath(filePath, testcase);
        System.out.println("Path"+docPathLocal);
        SearchMenuWM("WM Mobile","WMMobile");
        // SearchInWMMobile("JD Pack Pick Cart");//", "jdpackpickcart");
        Thread.sleep(5000);
        SearchInWmMobile("JD ECOM Pack Pick Cart","jdecompackpickcart");
        //Whatever present in Column E it will pack
        EcomPacking(filePath,testcase);

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



    public static void EcomPacking(String filePath, String testcaseName) {
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




            String token = getAuthTokenFromExcel();

            for (String olpn : olpns) {
                System.out.println("Packing OLPN: " + olpn);

                String responseBody = callOlpnApi(token, olpn);
                System.out.println("Packing OLPN: " + responseBody);
                List<JSONObject> completedResults = getCompletedResults(responseBody);
                System.out.println("Will execute the transaction");
                runScanning(completedResults);





                try {
                    WebDriverWait shortWait = new WebDriverWait(driver, Duration.ofSeconds(time));
                    WebElement okButton = shortWait.until(
                            ExpectedConditions.elementToBeClickable(By.xpath("//button[.//span[text()='Ok']]"))
                    );
                    js.executeScript("arguments[0].scrollIntoView(true);", okButton);
                    Thread.sleep(3000);
                    captureScreenshot("Click OK ");
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
    public static String getAuthTokenFromExcel() throws IOException {
        ExcelReader reader = new ExcelReader();

// By header name
        String LOGIN_URL = reader.getCellValueByHeader(1, "LOGIN_URL");
        String UIUsername = reader.getCellValueByHeader(1, "username");
        String UIPassword = reader.getCellValueByHeader(1, "password");

        reader.close();


        // Step 2: Call token API
        OkHttpClient client = new OkHttpClient();
        MediaType mediaType = MediaType.parse("application/x-www-form-urlencoded");
        RequestBody body = RequestBody.create(mediaType,
                "grant_type=password&username=" + UIUsername + "&password=" + UIPassword);

        Request request = new Request.Builder()
                .url(LOGIN_URL)
                .method("POST", body)
                .addHeader("Content-Type", "application/x-www-form-urlencoded")
                .addHeader("Authorization", "Basic dWpkc3N0YWdlMTpFYXJ0aC1Nb29uLVN1bjE=")
                .build();

        Response response = client.newCall(request).execute();
        String responseBody = response.body() != null ? response.body().string() : null;
        JsonObject json = JsonParser.parseString(responseBody).getAsJsonObject();

        return json.has("access_token") ? json.get("access_token").getAsString() : null;
    }









    private static final OkHttpClient client = new OkHttpClient();

    // --- Method to call OLPN API ---
    public static String callOlpnApi(String token, String olpn) throws Exception {
        ExcelReader reader = new ExcelReader();
        String BASE_URL = reader.getCellValueByHeader(1, "BASE_URL");
        String SelectedOrganization = reader.getCellValueByHeader(1, "SelectedOrganization");
        String SelectedLocation = reader.getCellValueByHeader(1, "SelectedLocation");
        reader.close();

        JSONObject filter = new JSONObject()
                .put("ViewName", "taskDetail")
                .put("AttributeId", "OlpnId")
                .put("DataType", JSONObject.NULL)
                .put("requiredFilter", false)
                .put("Operator", "=")
                .put("SupportsExactMatch", false)
                .put("FilterValues", new JSONArray().put(olpn))
                .put("negativeFilter", false);

        JSONObject body = new JSONObject()
                .put("ViewName", "TaskDetail")
                .put("Filters", new JSONArray().put(filter))
                .put("RequestAttributeIds", new JSONArray())
                .put("SearchOptions", new JSONArray())
                .put("SearchChains", new JSONArray())
                .put("FilterExpression", JSONObject.NULL)
                .put("Page", 0)
                .put("TotalCount", -1)
                .put("SortOrder", "desc")
                .put("MultiSort", new JSONArray())
                .put("SortIndicator", "chevron-up")
                .put("TimeZone", "Europe/Paris")
                .put("IsCommonUI", false)
                .put("ComponentShortName", JSONObject.NULL)
                .put("EnableMaxCountLimit", true)
                .put("MaxCountLimit", 500)
                .put("PageQuery", JSONObject.NULL)
                .put("ChildQuery", JSONObject.NULL)
                .put("ComponentName", "com-manh-cp-task")
                .put("Size", 25)
                .put("AdvancedFilter", false)
                .put("Sort", "CreatedTimestamp");

        MediaType mediaType = MediaType.parse("application/json");
        RequestBody requestBody = RequestBody.create(body.toString(), mediaType);

        Request request = new Request.Builder()
                .url(BASE_URL + "/dmui-facade/api/dmui-facade/entity/search")   // replace with actual endpoint
                .method("POST", requestBody)
                .addHeader("Content-Type", "application/json")
                .addHeader("SelectedOrganization", SelectedOrganization)
                .addHeader("SelectedLocation", SelectedLocation)
                .addHeader("Authorization", "Bearer " + token)
                .build();

        try (Response response = client.newCall(request).execute()) {
            if (response.body() == null) {
                return null;
            }
            return response.body().string();
        }
    }

    // --- Method to parse response and return Completed results sorted by OlpnDetailId ---
    public static List<JSONObject> getCompletedResults(String responseBody) {
        List<JSONObject> completed = new ArrayList<>();
        if (responseBody == null || responseBody.isEmpty()) return completed;

        JSONObject jsonResponse = new JSONObject(responseBody);
        JSONArray results = jsonResponse.getJSONObject("data").getJSONArray("Results");

        for (int i = 0; i < results.length(); i++) {
            JSONObject r = results.getJSONObject(i);
            if ("Completed".equalsIgnoreCase(r.getString("DetailStatusDescription"))) {
                completed.add(r);
            }
        }

        completed.sort(Comparator.comparingInt(r -> Integer.parseInt(r.getString("OlpnDetailId"))));
        return completed;
    }

    // --- Method to run Selenium scanning logic ---
    public static void runScanning( List<JSONObject> completedResults) throws IOException, InterruptedException {
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(time));
        if (completedResults.isEmpty()) return;

        String tote = completedResults.get(0).getString("TargetContainerId");
        boolean allSameTote = completedResults.stream()
                .allMatch(r -> tote.equals(r.getString("TargetContainerId")));

        if (allSameTote) {
            System.out.println("Tote to be enter: " + tote);
            WebElement toteInput = wait.until(ExpectedConditions.visibilityOfElementLocated(
                    By.cssSelector("input[data-component-id='accepttote_barcodetextfield_tote']")));
            toteInput.clear();
            toteInput.sendKeys(tote);
            Thread.sleep(3000);
            captureScreenshot("Enter Tote ");
            DocPathManager.saveSharedDocument();
            System.out.println("Enter Tote screenshot done");
            Thread.sleep(3000);
            toteInput.sendKeys(Keys.ENTER);
        }
        System.out.println("Tote enter: " + tote);

        for (JSONObject r : completedResults) {
            String itemId = r.getString("ItemId");
            String olpnId = r.getString("OlpnId");
            System.out.println("Item to be enter: " + itemId);
            String token = getAuthTokenFromExcel();
            String itemCleaned = searchItemMaster(itemId,token);
            WebElement itemInput = wait.until(ExpectedConditions.visibilityOfElementLocated(
                    By.cssSelector("input[data-component-id='acceptitem_barcodetextfield_item']")));
            itemInput.clear();
            itemInput.sendKeys(itemCleaned);
            itemInput.sendKeys(Keys.ENTER);

            System.out.println("olpnId to be enter: " + olpnId);
            WebElement olpnInput = wait.until(ExpectedConditions.visibilityOfElementLocated(
                    By.cssSelector("input[data-component-id='verifyolpn_barcodetextfield_scanolpn']")));
            olpnInput.clear();
            olpnInput.sendKeys(olpnId);
            olpnInput.sendKeys(Keys.ENTER);
        }

        System.out.println("Done");
//            // Wait until the button is visible
//            WebElement closeButton = wait.until(
//                    ExpectedConditions.elementToBeClickable(
//                            By.cssSelector("button[data-component-id='action_closecontainer_button']")
//                    )
//            );
//
//// Click the button
//            closeButton.click();

    }


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



}