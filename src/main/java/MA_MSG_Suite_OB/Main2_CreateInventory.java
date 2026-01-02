

package MA_MSG_Suite_OB;

import com.google.gson.JsonArray;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;
import io.github.bonigarcia.wdm.WebDriverManager;
import org.openqa.selenium.WebDriver;
import okhttp3.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.*;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;


import okhttp3.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;


import java.io.*;

import java.util.*;




public class Main2_CreateInventory {

    public static WebDriver driver;
    public static int time =60;

    /**
     * Executes one row as a testcase. Marks Result column as Passed if all steps succeed,
     * otherwise Failed.
     */
//    public static void execute(Row row, String filePath, String sheetName, String env) {
//        boolean isPassed = false;
//
//        try {
//            String item = getString(row, 1);
//            String ilpnId = getString(row, 2);
//            String itemBarcode = getString(row, 3);
//            int quantity = getInt(row, 4);
//
//            if (item.isEmpty() || ilpnId.isEmpty() || itemBarcode.isEmpty() || quantity == 0) {
//                throw new IllegalArgumentException("Skipping row due to missing or invalid data.");
//            }
//
//            System.out.printf("Inventory processing: Item=%s, ILPN=%s, Barcode=%s, Quantity=%d%n",
//                    item, ilpnId, itemBarcode, quantity);
//
//
//
//
//
//
//
//            String token = getAuthTokenFromExcel();
//            if (token == null) throw new RuntimeException("Failed to retrieve access token.");
//
//            String jsonBody = buildCreateIlpnPayload(item, ilpnId, itemBarcode, quantity);
//            String response = callCreateIlpnAPI(jsonBody, token);
//            System.out.println("Create ILPN Response:\n" + response);
//
//            WebDriverManager.chromedriver().setup();
//            ChromeOptions options = new ChromeOptions();
//            options.addArguments("--start-maximized");
//            driver = new ChromeDriver(options);
//
//            Main1_URL_Login1 login1 = new Main1_URL_Login1(driver, env);
//            login1.execute();
//            System.out.println("Login done.");
//
//            SearchandOpenWMMobie();
//            SearchInWMMobie("JD Carton Putaway");
//            isPassed = putaway(ilpnId);
//
//
//            Thread.sleep(15000);
//
//
//        } catch (Exception e) {
//            System.err.println("❌ Testcase failed: " + e.getMessage());
//            e.printStackTrace(System.err);
//        } finally {
//            closeExcelIfOpen();
//            updateResult(filePath, sheetName, row.getRowNum(), isPassed ? "Passed" : "Failed");
//            if (driver != null) {
//                driver.quit();
//            }
//        }
//    }











    public static void execute(String testcaseId, String filePath, String sheetName, String env) {
        boolean isPassed = false;
        //WebDriver driver1 = null;

        try {
            FileInputStream fis = new FileInputStream(filePath);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheet(sheetName);

            List<Row> matchingRows = new ArrayList<>();
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                String tc = getString(row, 0); // assuming column 0 is Testcase
                if (tc.equalsIgnoreCase(testcaseId)) {
                    matchingRows.add(row);
                }
            }

            if (matchingRows.isEmpty()) {
                System.out.println("No rows found for testcase: " + testcaseId);
                return;
            }

            String token = getAuthTokenFromExcel();
            if (token == null) throw new RuntimeException("Failed to retrieve access token.");


            for (Row row : matchingRows) {
                try {
                    String item = getString(row, 1);
                    String ilpnId = getString(row, 2);
                    String itemBarcode = getString(row, 3);
                    int quantity = getInt(row, 4);
                    String locationBarcode = getString(row, 5);

                    if (item.isEmpty() || ilpnId.isEmpty() || itemBarcode.isEmpty() || quantity == 0 || locationBarcode.isEmpty()) {
                        throw new IllegalArgumentException("Skipping row due to missing or invalid data.");
                    }

                    System.out.printf(
                            "Inventory processing: Item=%s, ILPN=%s, Barcode=%s, Quantity=%d, locationBarcode=%s%n",
                            item, ilpnId, itemBarcode, quantity, locationBarcode
                    );

                    String jsonBody = buildCreateIlpnPayload(item, ilpnId, itemBarcode, quantity);
                    String response = callCreateIlpnAPI(jsonBody, token);
                    System.out.println("Create ILPN Response:\n" + response);

                    WebDriverManager.chromedriver().setup();
                    ChromeOptions options = new ChromeOptions();
                    options.addArguments("--start-maximized");
                    driver = new ChromeDriver(options);

                    Main1_URL_Login1 login1 = new Main1_URL_Login1(driver, env);
                    login1.execute();
                    System.out.println("Login done.");
                    Thread.sleep(5000);

                    SearchMenuWM("WM Mobile","WMMobile");

                    SearchInWmMobile("JD Carton Putaway", "jdcartonputaway");

                    isPassed = putaway(ilpnId);

                   // isPassed=true;

                    workbook.close();
                    fis.close();

                } catch (Exception e) {
                    System.err.println("❌ Row failed: " + e.getMessage());
                    e.printStackTrace(System.err);
                } finally {
                    closeExcelIfOpen();
                    updateResult(filePath, sheetName, row.getRowNum(), isPassed ? "Passed" : "Failed");
                }
            }



        } catch (Exception e) {
            System.err.println("❌ Testcase execution failed: " + e.getMessage());
            e.printStackTrace(System.err);
        } finally {

            if (driver != null) {
                driver.quit();
            }
        }
    }

    /**
     * Updates the Result column for a given row index
     */
    private static void updateResult(String filePath, String sheetName, int rowIndex, String status) {
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheet(sheetName);
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                Cell resultCell = row.getCell(6); // assuming Result is column 6
                if (resultCell == null) resultCell = row.createCell(6, CellType.STRING);
                resultCell.setCellValue(status);
            }

            try (FileOutputStream fos = new FileOutputStream(filePath)) {
                workbook.write(fos);
            }
            System.out.println("Result updated for row " + rowIndex + ": " + status);

        } catch (IOException e) {
            System.err.println("❌ Failed to update Excel: " + e.getMessage());
        }
    }

    public static String getAuthTokenFromExcel() throws IOException {
        ExcelReader reader = new ExcelReader();

// By header name
        String LOGIN_URL = reader.getCellValueByHeader(1, "LOGIN_URL");
        String UIUsername = reader.getCellValueByHeader(1, "username");
        String UIPassword = reader.getCellValueByHeader(1, "password");

        reader.close();

        // System.err.println("❌ Testcase failed: " + LOGIN_URL+" ,"+UIPassword+ "  ,"+UIUsername);
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

    private static String getString(Row row, int index) {
        Cell cell = row.getCell(index);
        return cell != null ? cell.toString().trim() : "";
    }

    private static void closeExcelIfOpen() {
        try {
            Process process = Runtime.getRuntime().exec("tasklist");
            BufferedReader reader = new BufferedReader(new InputStreamReader(process.getInputStream()));
            String line;
            boolean excelRunning = false;
            while ((line = reader.readLine()) != null) {
                if (line.toLowerCase().contains("excel.exe")) {
                    excelRunning = true;
                    break;
                }
            }
            if (excelRunning) {
                System.out.println("⚠️ Excel is open. Closing it...");
                Runtime.getRuntime().exec("taskkill /IM excel.exe /F");
                Thread.sleep(2000);
            }
        } catch (Exception e) {
            System.err.println("⚠️ Could not check/close Excel: " + e.getMessage());
        }
    }
    private static String buildCreateIlpnPayload(String item, String ilpnId, String barcode, int quantity) {
        JsonObject root = new JsonObject();
        root.add("AdditionalFields", new JsonObject());
        root.addProperty("IlpnId", ilpnId);
        root.addProperty("ReasonCode", "IA");
        root.addProperty("SkipLabelPrinting", true);
        root.addProperty("TransactionId", "SeedCreateIlpnId");

        JsonObject taskIntegration = new JsonObject();
        taskIntegration.addProperty("LaborActivityId", "Create ilpn");
        taskIntegration.addProperty("TransactionId", "SeedCreateIlpnId");
        taskIntegration.addProperty("TransactionTypeId", "Create iLPN");
        taskIntegration.addProperty("WorkflowInitTime", "2019-08-24T14:15:22Z");
        root.add("TaskIntegrationDTO", taskIntegration);

        JsonArray scannedInventory = new JsonArray();
        JsonObject inventoryAttributes = new JsonObject();
        inventoryAttributes.addProperty("Item", item);
        inventoryAttributes.addProperty("ItemBarcode", barcode);
        inventoryAttributes.addProperty("CompareAttributes", true);
        inventoryAttributes.addProperty("TrackInventoryType", true);

        JsonObject scannedItem = new JsonObject();
        scannedItem.add("InventoryAttributes", inventoryAttributes);
        scannedItem.addProperty("ScannedQuantity", quantity);
        scannedInventory.add(scannedItem);

        root.add("ScannedInventory", scannedInventory);
        return root.toString();
    }
    public static String callCreateIlpnAPI(String jsonBody, String token) throws IOException {
        ExcelReader reader = new ExcelReader();

// By header name
        String BASE_URL = reader.getCellValueByHeader(1, "BASE_URL");
        String SelectedOrganization = reader.getCellValueByHeader(1, "SelectedOrganization");
        String SelectedLocation = reader.getCellValueByHeader(1, "SelectedLocation");

        reader.close();






        OkHttpClient client = new OkHttpClient();
        MediaType mediaType = MediaType.parse("application/json");
        RequestBody requestBody = RequestBody.create(mediaType, jsonBody);

        Request request = new Request.Builder()
                .url(BASE_URL + "/inventory-management/api/inventory-management/create/createIlpnAndStartTask")
                .method("POST", requestBody)
                .addHeader("Content-Type", "application/json")
                .addHeader("Authorization", "Bearer " + token)
                .addHeader("SelectedOrganization", SelectedOrganization)
                .addHeader("SelectedLocation", SelectedLocation)
                .build();

        Response response = client.newCall(request).execute();
        return response.body() != null ? response.body().string() : "No response body";
    }

    private static int getInt(Row row, int index) {
        Cell cell = row.getCell(index);
        if (cell == null) return 0;
        switch (cell.getCellType()) {
            case NUMERIC:
                return (int) cell.getNumericCellValue();
            case STRING:
                try {
                    return Integer.parseInt(cell.getStringCellValue().trim());
                } catch (NumberFormatException e) {
                    return 0;
                }
            case FORMULA:
                try {
                    return (int) cell.getNumericCellValue();
                } catch (Exception e) {
                    return 0;
                }
            default:
                return 0;
        }
    }
    public static WebElement findVisibleElement(WebDriver driver, By locator) {
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(time));
        return wait.until(ExpectedConditions.visibilityOfElementLocated(locator));
    }
    public static void SearchandOpenWMMobie() {
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(time));
        JavascriptExecutor js = (JavascriptExecutor) driver;

        try {
            // Step 1: Wait for the shadow host
            WebElement shadowHost = wait.until(
                    ExpectedConditions.presenceOfElementLocated(
                            By.cssSelector("ion-button[data-component-id='menu-toggle-button']")
                    )
            );

            // Step 2: Access shadow root
            SearchContext shadowRoot = (SearchContext) js.executeScript("return arguments[0].shadowRoot", shadowHost);

            // Step 3: Click native button inside shadow root
            WebElement nativeButton = shadowRoot.findElement(By.cssSelector("button.button-native"));
            nativeButton.click();
            System.out.println("Menu toggle button clicked.");

            // Step 4: Handle search input (could be ion-input, not plain input)
            WebElement searchInput;
            try {
                searchInput = wait.until(ExpectedConditions.presenceOfElementLocated(
                        By.cssSelector("ion-input[placeholder='Search Menu...']")
                ));
            } catch (TimeoutException e) {
                // fallback to plain input
                searchInput = wait.until(ExpectedConditions.presenceOfElementLocated(
                        By.xpath("//input[@placeholder='Search Menu...']")
                ));
            }

            // If it's an ion-input, you may need to go into its shadow root
            try {
                SearchContext inputShadow = (SearchContext) js.executeScript("return arguments[0].shadowRoot", searchInput);
                WebElement innerInput = inputShadow.findElement(By.cssSelector("input"));
                innerInput.clear();
                innerInput.sendKeys("WM Mobile");
            } catch (Exception e) {
                // fallback if it's a normal input
                searchInput.clear();
                searchInput.sendKeys("WM Mobile");
            }
            System.out.println("Search Done WM Mobile");

            // Step 5: Click WM Mobile button
            WebElement wmMobileButton = wait.until(ExpectedConditions.presenceOfElementLocated(
                    By.cssSelector("button#wmMobile[data-component-id='WMMobile']")
            ));
            js.executeScript("arguments[0].scrollIntoView(true);", wmMobileButton);
            js.executeScript("arguments[0].click();", wmMobileButton);
            System.out.println("✅ 'WM Mobile' button clicked.");

            // Step 6: Switch to new tab
            Thread.sleep(3000);
            ArrayList<String> tabs = new ArrayList<>(driver.getWindowHandles());
            if (tabs.size() > 1) {
                driver.switchTo().window(tabs.get(1));
            }
            Thread.sleep(3000);

        } catch (Exception e) {
            System.err.println("❌ Error in SearchandOpenWMMobie: " + e.getMessage());
            e.printStackTrace(System.err);
        }
    }
    public static void SearchInWMMobie(String Input){
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(time));
        JavascriptExecutor js = (JavascriptExecutor) driver;
        try{
            WebElement searchInput1 = wait.until(ExpectedConditions.elementToBeClickable(
                    By.cssSelector("input.searchbar-input[placeholder='Search']")));
            By.xpath("//input[@type='search' and @placeholder='Search']");
            searchInput1.click();
            searchInput1.clear();
            Thread.sleep(3000);
            searchInput1.sendKeys( Input);
        } catch (Exception e) {
            System.err.println("❌ Error in search " + e.getMessage());
            e.printStackTrace();
        }}
    public static boolean putaway(String lpn) {
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(time));
        JavascriptExecutor js = (JavascriptExecutor) driver;

        try {
//        WebElement element = driver.findElement(By.cssSelector("ion-label[data-component-id='jdcartonputaway']"));
//        js.executeScript("arguments[0].scrollIntoView(true);", element);
//        element.click();
//        System.out.println("✅ Clicked on 'User Directed Putaway'");

            WebElement containerInput = findVisibleElement(
                    driver,
                    By.cssSelector("input[data-component-id='acceptcontainerforsystemdirectedputaway_barcodetextfield_container']")

            );
            containerInput.sendKeys(lpn, Keys.ENTER);

            // 🔍 Check for error popup after entering container
            try {
                WebElement errorPopup = wait.until(ExpectedConditions.presenceOfElementLocated(
                        By.cssSelector("div[data-component-id='popover_error_message']")
                ));
                String errorText = errorPopup.getText();
                System.err.println("❌ Error popup detected: " + errorText);
                return false; // mark testcase as failed
            } catch (TimeoutException noPopup) {
                // no popup appeared, continue normal flow
            }

            Thread.sleep(8000);

            WebElement locationElement = findVisibleElement(
                    driver,
                    By.cssSelector("ion-col[data-component-id='acceptlocationforsystemdirectedputaway_barcodetextfield_location']")

            );
            String locationID = (String) js.executeScript("return arguments[0].innerText;", locationElement);
            String locationBarcode1 = locationID.replaceAll("-", "");
            System.out.println("✅ Replace completed for ILPN: " + locationBarcode1);

            WebElement locationInput = findVisibleElement(
                    driver,
                    By.cssSelector("input[data-component-id='acceptlocationforsystemdirectedputaway_barcodetextfield_scanlocation']")
            );
            js.executeScript("arguments[0].value = arguments[1];", locationInput, locationBarcode1);
            locationInput.sendKeys(Keys.ENTER);

            Thread.sleep(5000);
            System.out.println("✅ Putaway completed for ILPN: " + lpn);
            return true; // success

        } catch (Exception e) {
            System.err.println("❌ Putaway failed for ILPN " + lpn + ": " + e.getMessage());
            return false;
        }
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


        } catch (Exception e) {
            System.err.println("❌ Error in "+Transaction + e.getMessage());
            e.printStackTrace();
        }
        // Locate the ion-label using its data-component-id


        WebElement labelElement = driver.findElement(
                By.cssSelector("ion-label[data-component-id="+ComponentId+"]")
        );


        // Scroll into view to ensure it's interactable
        ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", labelElement);

        // Click using JavaScript (in case native click doesn't work)
        ((JavascriptExecutor) driver).executeScript("arguments[0].click();", labelElement);

        System.out.println("Clicked on" +Transaction+" label.");


    }


}
