package MA_MSG_Suite_OB;

import com.google.gson.*;
import com.google.gson.stream.JsonReader;
import okhttp3.*;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

import com.google.gson.JsonArray;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;
import io.github.bonigarcia.wdm.WebDriverManager;
import okhttp3.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.FileInputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.ArrayList;

public class UserDirectedPutaway {
    public static WebDriver driver;
    public static String getAuthToken() throws IOException {
        System.out.println("Fetching auth token...");
        OkHttpClient client = new OkHttpClient();
        MediaType mediaType = MediaType.parse("application/x-www-form-urlencoded");
        RequestBody body = RequestBody.create(mediaType, "grant_type=password&username=cogs&password=Cogs@123456");

        Request request = new Request.Builder()
                .url("https://ujdss-auth.sce.manh.com/oauth/token")
                .method("POST", body)
                .addHeader("Content-Type", "application/x-www-form-urlencoded")
                .addHeader("Authorization", "Basic dWpkc3N0YWdlMTpFYXJ0aC1Nb29uLVN1bjE=")
                .build();

        Response response = client.newCall(request).execute();
        String responseBody = response.body() != null ? response.body().string() : null;

        System.out.println("Auth response: " + responseBody);

        JsonObject json = JsonParser.parseString(responseBody).getAsJsonObject();
        String token = json.has("access_token") ? json.get("access_token").getAsString() : null;

        if (token == null) {
            System.err.println("Failed to retrieve access token.");
        } else {
            System.out.println("Access token retrieved successfully.");
        }

        return token;
    }

    public static String callCreateIlpnAPI(String jsonBody, String token) throws IOException {
        System.out.println("Sending API request...");
        OkHttpClient client = new OkHttpClient();
        MediaType mediaType = MediaType.parse("application/json");
        RequestBody requestBody = RequestBody.create(mediaType, jsonBody);

        Request request = new Request.Builder()
                .url("https://ujdss.sce.manh.com/inventory-management/api/inventory-management/create/createIlpnAndStartTask")
                .method("POST", requestBody)
                .addHeader("Content-Type", "application/json")
                .addHeader("Authorization", "Bearer " + token)
                .addHeader("SelectedOrganization", "HEERLEN51")
                .addHeader("SelectedLocation", "HEERLEN51")
                .build();

        Response response = client.newCall(request).execute();
        String responseText = response.body() != null ? response.body().string() : "No response body";

        System.out.println("API Response:\n" + responseText);
        return responseText;
    }

    public static void main(String[] args) throws IOException {
        String token = getAuthToken();
        if (token == null) return;

        System.out.println("Reading Excel file...");
        FileInputStream fis = new FileInputStream("C:\\Users\\2378594\\IdeaProjects\\Testcases - Copy\\Testcases - Copy\\OOdata.xlsx");
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheet("Data Creation");

        if (sheet == null) {
            System.err.println("Sheet 'Data Creation' not found.");
            workbook.close();
            fis.close();
            return;
        }

        System.out.println("Sheet found: " + sheet.getSheetName());
        System.out.println("Total rows: " + sheet.getLastRowNum());

        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) continue;

            String item = getString(row, 0);
            String ilpnId = getString(row, 1);
            String barcode = getString(row, 2);
            int quantity = getInt(row, 3);

            System.out.printf("Row %d: Item=%s, ILPN=%s, Barcode=%s, Quantity=%d%n", i, item, ilpnId, barcode, quantity);

            if (item.isEmpty() || ilpnId.isEmpty() || quantity == 0) {
                System.out.println("Skipping row due to missing or invalid data.");
                continue;
            }

            JsonObject root = new JsonObject();
            root.add("AdditionalFields", new JsonObject());
            root.addProperty("ConditionCodeId", "");
            root.addProperty("CriteriaId", "");
            root.addProperty("GeneratedIlpn", false);
            root.addProperty("IlpnExist", false);
            root.addProperty("IlpnId", ilpnId);
            root.addProperty("LpnSizeTypeId", "");
            root.addProperty("ReasonCode", "IA");
            root.addProperty("ReferenceCode", "");
            root.addProperty("SecondaryReferenceCode", "");
            root.addProperty("SkipLabelPrinting", true);
            root.addProperty("SourceContainerId", "");
            root.addProperty("SourceLocationId", "");
            root.addProperty("TaskId", "");

            JsonObject taskIntegration = new JsonObject();
            taskIntegration.addProperty("LaborActivityId", "Create ilpn");
            taskIntegration.addProperty("TaskId", "");
            taskIntegration.addProperty("TransactionId", "SeedCreateIlpnId");
            taskIntegration.addProperty("TransactionTypeId", "Create iLPN");
            taskIntegration.addProperty("WorkflowInitTime", "2019-08-24T14:15:22Z");
            root.add("TaskIntegrationDTO", taskIntegration);
            root.addProperty("TransactionId", "SeedCreateIlpnId");

            JsonArray scannedInventory = new JsonArray();
            JsonObject inventoryAttributes = new JsonObject();
            inventoryAttributes.add("AdditionalFields", new JsonObject());
            inventoryAttributes.addProperty("BatchNumber", "");
            inventoryAttributes.addProperty("BusinessUnitId", "");
            inventoryAttributes.addProperty("CompareAttributes", true);
            inventoryAttributes.addProperty("CountryOfOrigin", "");
            inventoryAttributes.addProperty("IncubationDateTypeId", "");
            inventoryAttributes.addProperty("IncubationDays", 0);
            inventoryAttributes.addProperty("IncubationHours", 0);
            inventoryAttributes.addProperty("InventoryAttribute1", "");
            inventoryAttributes.addProperty("InventoryAttribute2", "");
            inventoryAttributes.addProperty("InventoryAttribute3", "");
            inventoryAttributes.addProperty("InventoryAttribute4", "");
            inventoryAttributes.addProperty("InventoryAttribute5", "");
            inventoryAttributes.addProperty("InventoryTypeId", "");
            inventoryAttributes.addProperty("Item", item);
            inventoryAttributes.addProperty("ItemBarcode", barcode);
            inventoryAttributes.addProperty("ItemDescription", "");
            inventoryAttributes.addProperty("PackUomQuantity", 0);
            inventoryAttributes.addProperty("PackUomTypeId", "");
            inventoryAttributes.addProperty("ProductStatusId", "");
            inventoryAttributes.addProperty("TrackBatchNumber", true);
            inventoryAttributes.addProperty("TrackCountryOfOrigin", true);
            inventoryAttributes.addProperty("TrackExpiryDate", true);
            inventoryAttributes.addProperty("TrackInventoryAttribute1", true);
            inventoryAttributes.addProperty("TrackInventoryAttribute2", true);
            inventoryAttributes.addProperty("TrackInventoryAttribute3", true);
            inventoryAttributes.addProperty("TrackInventoryAttribute4", true);
            inventoryAttributes.addProperty("TrackInventoryAttribute5", true);
            inventoryAttributes.addProperty("TrackInventoryType", true);
            inventoryAttributes.addProperty("TrackManufacturingDate", true);
            inventoryAttributes.addProperty("TrackPackQuantity", true);
            inventoryAttributes.addProperty("TrackProductStatus", true);
            inventoryAttributes.addProperty("TrackShipByDate", true);

            JsonObject scannedItem = new JsonObject();
            scannedItem.add("InventoryAttributes", inventoryAttributes);
            scannedItem.addProperty("ScannedQuantity", quantity);
            scannedInventory.add(scannedItem);

            root.add("ScannedInventory", scannedInventory);

            String jsonBody = root.toString();
            System.out.println("Request JSON:\n" + jsonBody);

            String response = callCreateIlpnAPI(jsonBody, token);
            System.out.println("Response for ILPN " + ilpnId + ":\n" + response);
        }

        workbook.close();
        fis.close();


        login(); // Call login only once per unique OPS

        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) continue;

            String item = getString(row, 0);
            String ilpnId = getString(row, 1);
            String barcode = getString(row, 2);
            int quantity = getInt(row, 3);
            String locationBarcode = getString(row, 4);

            System.out.printf("Row %d: Item=%s, ILPN=%s, Barcode=%s, Quantity=%d, Location=%s%n",
                    i, item, ilpnId, barcode, quantity, locationBarcode);

            if (item.isEmpty() || ilpnId.isEmpty() || barcode.isEmpty() || quantity == 0 || locationBarcode.isEmpty()) {
                System.out.println("Skipping row due to missing or invalid data.");
                continue;
            }

            // ✅ Call putaway for each row
            putaway(item, ilpnId, barcode, quantity, locationBarcode);
        }



    }

    private static String getString(Row row, int index) {
        Cell cell = row.getCell(index);
        return cell != null ? cell.toString().trim() : "";
    }

    private static int getInt(Row row, int index) {
        Cell cell = row.getCell(index);
        if (cell != null && cell.getCellType() == CellType.NUMERIC) {
            return (int) cell.getNumericCellValue();
        } else if (cell != null && cell.getCellType() == CellType.STRING) {
            try {
                return Integer.parseInt(cell.getStringCellValue().trim());
            } catch (NumberFormatException e) {
                return 0;
            }
        }
        return 0;
    }

    public static void login()
    {
        // WebDriverManager (if internet access is available)
        WebDriverManager.chromedriver().setup();
        driver = new ChromeDriver();

        driver.get("https://ujdss-auth.sce.manh.com/auth/realms/maactive/protocol/openid-connect/auth?scope=openid&client_id=zuulserver.1.0.0&redirect_uri=https://ujdss.sce.manh.com/login&response_type=code&state=52FrgC");

        driver.manage().window().maximize();


        WebElement usernameField = driver.findElement(By.id("username"));
        WebElement passwordField = driver.findElement(By.id("password"));

        usernameField.sendKeys("cogs");
        passwordField.sendKeys("Cogs@123456");


        driver.findElement(By.id("kc-login")).click();
        try {
            Thread.sleep(20000); // 20,000 milliseconds = 20 seconds
        } catch (InterruptedException e) {
            e.printStackTrace();
        }


    }
    public static void putaway(String item, String lpn, String barcodeValue, int quantity, String locationBarcode) {
        System.out.println("📦 Starting Putaway for ILPN: " + lpn);

        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
        JavascriptExecutor js = (JavascriptExecutor) driver;

        try {
            // Step 1: Navigate to 'User Directed Putaway'
            WebElement shadowHost = driver.findElement(By.cssSelector("ion-button.menu-toggle-button"));
            SearchContext shadowRoot = (SearchContext) js.executeScript("return arguments[0].shadowRoot", shadowHost);
            WebElement menuButton = shadowRoot.findElement(By.cssSelector("button.button-native"));
            menuButton.click();

            wait.until(ExpectedConditions.invisibilityOfElementLocated(
                    By.cssSelector("ion-popover.workflow-error-popover")));

            WebElement searchInput = wait.until(ExpectedConditions.elementToBeClickable(
                    By.xpath("//input[@placeholder='Search Menu...']")));
            searchInput.clear();
            searchInput.sendKeys("WM Mobile");

            WebElement wmMobileButton = wait.until(ExpectedConditions.presenceOfElementLocated(
                    By.cssSelector("button#wmMobile[data-component-id='WMMobile']")));
            js.executeScript("arguments[0].scrollIntoView(true);", wmMobileButton);
            js.executeScript("arguments[0].click();", wmMobileButton);

            Thread.sleep(3000);
            ArrayList<String> tabs = new ArrayList<>(driver.getWindowHandles());
            driver.switchTo().window(tabs.get(tabs.size() - 1));
            Thread.sleep(3000);

            WebElement searchInput1 = wait.until(ExpectedConditions.elementToBeClickable(
                    By.cssSelector("input.searchbar-input[placeholder='Search']")));
            searchInput1.click();
            searchInput1.clear();
            Thread.sleep(3000);
            searchInput1.sendKeys("User Directed Putaway");

            WebElement putawayButtonHost = wait.until(ExpectedConditions.presenceOfElementLocated(
                    By.cssSelector("ion-item[data-component-id='userdirectedputaway']")));
            SearchContext putawayShadowRoot = (SearchContext) js.executeScript("return arguments[0].shadowRoot", putawayButtonHost);
            WebElement clickableDiv = putawayShadowRoot.findElement(By.cssSelector("div.item-native"));
            js.executeScript("arguments[0].scrollIntoView(true);", clickableDiv);
            js.executeScript("arguments[0].click();", clickableDiv);

            // Step 2: Enter ILPN
            WebElement containerInput = wait.until(ExpectedConditions.visibilityOfElementLocated(
                    By.cssSelector("input[data-component-id='acceptcontainerforuserdirectedputaway_barcodetextfield_container']")));
            containerInput.click();
            containerInput.clear();
            containerInput.sendKeys(lpn);
            containerInput.sendKeys(Keys.ENTER);

            // Step 3: Enter Barcode
//            WebElement scanContainerInput = wait.until(ExpectedConditions.presenceOfElementLocated(
//                    By.cssSelector("input[data-component-id='acceptcontainerforuserdirectedputaway_barcodetextfield_scancontainer']")));
//            js.executeScript("arguments[0].scrollIntoView(true);", scanContainerInput);
//            js.executeScript("arguments[0].value = arguments[1];", scanContainerInput, barcodeValue);
//            js.executeScript("arguments[0].dispatchEvent(new Event('input', { bubbles: true }));", scanContainerInput);
//            scanContainerInput.sendKeys(Keys.ENTER);

            // Step 4: Enter Quantity
            WebElement quantityInput = wait.until(ExpectedConditions.presenceOfElementLocated(
                    By.cssSelector("input[data-component-id='acceptitemquantity_naturalquantityfield_oversize-ex01']")));
            js.executeScript("arguments[0].scrollIntoView(true);", quantityInput);
            js.executeScript("arguments[0].value = arguments[1];", quantityInput, String.valueOf(quantity));
            js.executeScript("arguments[0].dispatchEvent(new Event('input', { bubbles: true }));", quantityInput);
            quantityInput.sendKeys(Keys.ENTER);


// ✅ Step 4: Enter Location Barcode (corrected selector)
            WebElement locationInput = wait.until(ExpectedConditions.presenceOfElementLocated(
                    By.cssSelector("input[data-component-id='acceptlocationforuserdirectedputaway_barcodetextfield_location']")));
            js.executeScript("arguments[0].scrollIntoView(true);", locationInput);
            js.executeScript("arguments[0].value = '';", locationInput); // Clear
            js.executeScript("arguments[0].value = arguments[1];", locationInput, locationBarcode);
            js.executeScript("arguments[0].dispatchEvent(new Event('input', { bubbles: true }));", locationInput);
            locationInput.sendKeys(Keys.ENTER);

            System.out.println("✅ Putaway completed for ILPN: " + lpn);

        } catch (Exception e) {
            System.err.println("❌ Error during putaway for ILPN " + lpn + ": " + e.getMessage());
            e.printStackTrace();
            // Optional: dump page source for debugging
            // System.out.println(driver.getPageSource());
        }

    }


}
