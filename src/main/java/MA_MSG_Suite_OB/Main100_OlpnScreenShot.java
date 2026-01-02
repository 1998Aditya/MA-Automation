package MA_MSG_Suite_OB;

import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;
import com.google.gson.stream.JsonReader;
import okhttp3.*;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import io.github.bonigarcia.wdm.WebDriverManager;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.StringReader;
import java.time.Duration;
import java.util.*;
import java.util.NoSuchElementException;
import java.util.stream.Collectors;

import static MA_MSG_Suite_OB.DocPathManager.captureAllCardsScreenshots;
import static MA_MSG_Suite_OB.DocPathManager.captureScreenshot;
import static java.lang.Thread.sleep;


public class Main100_OlpnScreenShot {

    //public static int time = 60;




        // ---------- Config ----------
        public static WebDriver driver;
        public static long time = 60;
        public static XWPFDocument document = new XWPFDocument();

    private static final OkHttpClient httpClient = new OkHttpClient.Builder()
            .followRedirects(true).followSslRedirects(true)
            .callTimeout(Duration.ofSeconds(time))
            .connectTimeout(Duration.ofSeconds(time))
            .readTimeout(Duration.ofSeconds(time))
            .writeTimeout(Duration.ofSeconds(time))
            .build();



//        public static String EXCEL_PATH = "C:\\Users\\2389120\\IdeaProjects\\msg-runner\\OOdata.xlsx";
//        public static String TESTCASE_VALUE = "TST_001";

        // Menu item (native <button>) details
        private static final String MENU_ITEM_TEXT   = "oLPNs 2.0";
        private static final String MENU_BUTTON_ID   = "olpnVer2";   // <-- from your DOM
        private static final String MENU_COMPONENT_ID= "oLPNs2.0";   // <-- from your DOM

        // Page component IDs
        private static final String COMP_ID_CHEVRON_UP   = "DMOlpnVer2-oLPN-chevron-up";
        private static final String COMP_ID_CHEVRON_DOWN = "DMOlpnVer2-oLPN-chevron-down"; // may or may not exist
        private static final String COMP_ID_OLPN_INPUT   = "OlpnId";

        private static final Duration DEFAULT_WAIT = Duration.ofSeconds(60);
        private static final Duration SHORT_WAIT   = Duration.ofSeconds(3);

    // ---------- Entry point ----------
      //  public static void main(String[] args) {
        public static void main(String EXCEL_PATH,String TESTCASE_VALUE,WebDriver driver,String env,String docPathLocal) {

            if (driver == null) {
                System.out.println("Driver is NULL33");
            } else {
                System.out.println("Driver is initialized");
            }



            WebDriver existingDriver = driver;
            Main100_OlpnScreenShot.driver = existingDriver;

            try {
                System.out.println("Olpn Screenshot Started for "+ TESTCASE_VALUE);
//                WebDriverManager.chromedriver().setup();
//                ChromeOptions options = new ChromeOptions();
//                options.addArguments("--start-maximized");
//
//                driver = new ChromeDriver(options);
//                driver.manage().window().maximize();
//                Main1_URL_Login1 login1 = new Main1_URL_Login1(driver, env);
//                login1.execute();
//                System.out.println("login done:");

                // Search the menu ONCE and navigate to oLPNs 2.0 page
                SearchMenuOnce(driver);

                // Fetch OLPNs from Excel by testcase
                String[] olpns = fetchOlpnsForDeclaredTestcase(EXCEL_PATH,TESTCASE_VALUE);
                if (olpns == null || olpns.length == 0) {
                    System.out.println("No OLPNs found for testcase: " + TESTCASE_VALUE);
                    return;
                }
                sleep(8000);
                System.out.println("\n=== Starting OLPN run for " + olpns.length + " items ===");
                for (int i = 0; i < olpns.length; i++) {
                    String olpn = olpns[i];

                    String token = getAuthTokenFromExcel();
                    if (token == null) {
                        System.err.println("❌ Failed to authenticate.");
                        return;
                    }

                    // Wait up to 300 attempts, 3 seconds each

                    waitUntilUcsPackingReady(olpn, token, 300, 3000);
                    System.out.println("UCSPacking is ready. Continue workflow...");





                    System.out.println("\n--- (" + (i + 1) + "/" + olpns.length + ") OLPN = " + olpn + " ---");
                    sortChevronAndEnterOlpnOnce(olpn,driver);
                    try { sleep(8000);
                        System.out.println("Waiting for screenshot");
                        //Add Screenshot Methods here
                        System.out.println("Output doc: " + docPathLocal);

                        sleep(10000);
                        captureScreenshot("Orders",driver);
                      //  captureAllCardsScreenshots();

                        DocPathManager.saveSharedDocument(); // optional: save after each class
                        System.out.println("\n✔ olpn Screenshot done");
                        selectCardAndOpenTask(driver);
                        System.out.println("\n✔ Task Screenshot done");

                    } catch (InterruptedException ignored) {}
                }

                System.out.println("\n✔ Completed OLPN flow for all items.");
            } catch (Exception e) {
                System.out.println("✖ Failure: " + e.getMessage());
                e.printStackTrace();
            } finally {
                try { sleep(500); } catch (InterruptedException ignored) {}
                if (driver != null) driver.quit();
            }
        }







    public static String getAuthTokenFromExcel() throws IOException {
        ExcelReader reader = new ExcelReader();
        String LOGIN_URL   = reader.getCellValueByHeader(1, "LOGIN_URL");
        String UIUsername  = reader.getCellValueByHeader(1, "username");
        String UIPassword  = reader.getCellValueByHeader(1, "password");
        reader.close();

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

    private static String readString(JsonObject obj, String key) {
        return (obj.has(key) && !obj.get(key).isJsonNull()) ? obj.get(key).getAsString().trim() : null;
    }
    public static Map<String, String> getUcsPackingStatusBatch(List<String> olpns, String bearerToken) throws IOException {

        Map<String, String> olpnStatusMap = new LinkedHashMap<>();

        if (olpns == null || olpns.isEmpty()) {
            System.err.println("⚠️ No OLPNs provided.");
            return olpnStatusMap;
        }

        String inList = olpns.stream()
                .filter(s -> s != null && !s.isBlank())
                .map(String::trim)
                .collect(Collectors.joining(","));

        String body = "{\"Query\":\"OlpnId in (" + inList + ")\"}";
       // System.out.println("📤 Batch body: " + body);

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

        try (Response response = httpClient.newCall(request).execute()) {

            String responseBody = response.body() != null ? response.body().string() : "";
            String contentType = Optional.ofNullable(response.header("Content-Type")).orElse("");

            if (response.code() >= 200 && response.code() < 300 &&
                    contentType.toLowerCase(Locale.ROOT).contains("application/json")) {
               // System.out.println("📤 Batch body: " + responseBody);
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

                            String ucsReady = "No";

                            if (item.has("Extended") && item.get("Extended").isJsonObject()) {
                                JsonObject ext = item.getAsJsonObject("Extended");

                                if (ext.has("MAUJDSReadyForUCSPacking") && !ext.get("MAUJDSReadyForUCSPacking").isJsonNull()) {
                                    boolean val = ext.get("MAUJDSReadyForUCSPacking").getAsBoolean();
                                    ucsReady = val ? "Yes" : "No";
                                }
                            }

                            // ✅ FIX: store in the correct map
                            olpnStatusMap.put(olpn, ucsReady);

                            System.out.println("↪️ OLPN " + olpn + " → MAUJDSReadyForUCSPacking: " + ucsReady);
                        }


                    }
                }
            }

        } catch (Exception e) {
            System.err.println("❌ Batch request failed: " + e);
        }

        return olpnStatusMap;
    }

    public static Map<String, List<String>> groupOlpnsByUcsPackingStatus(List<String> olpns, String bearerToken) throws IOException {

        Map<String, String> statusMap = getUcsPackingStatusBatch(olpns, bearerToken);
        Map<String, List<String>> grouped = new LinkedHashMap<>();

        for (String id : olpns) {
            String status = statusMap.getOrDefault(id, "No");
            grouped.computeIfAbsent(status, k -> new ArrayList<>()).add(id);
        }

        return grouped;
    }




    public static void waitUntilUcsPackingReady(String olpn, String bearerToken,
                                                int maxAttempts, long waitMillis) throws Exception {

        for (int attempt = 1; attempt <= maxAttempts; attempt++) {

            System.out.println("🔄 Attempt " + attempt + ": Checking UCSPacking status for OLPN " + olpn);

            Map<String, String> statusMap = getUcsPackingStatusBatch(Collections.singletonList(olpn), bearerToken);
            String status = statusMap.getOrDefault(olpn, "No");

            System.out.println("   → Status: " + status);

            if ("Yes".equalsIgnoreCase(status)) {
                System.out.println("✅ UCSPacking is READY for OLPN " + olpn);
                return;
            }

            Thread.sleep(waitMillis);
        }

        throw new RuntimeException("⏳ Timeout: UCSPacking never became READY for OLPN " + olpn);
    }













    // ---------- Search menu ONCE and navigate to oLPNs 2.0 ----------
        private static void SearchMenuOnce(WebDriver driver) throws InterruptedException {
            WebDriverWait wait1 = new WebDriverWait(driver, Duration.ofSeconds(time));
            WebDriverWait wait  = new WebDriverWait(driver, Duration.ofSeconds(time));
            JavascriptExecutor js = (JavascriptExecutor) driver;

            // 1) Toggle the side menu (ion-button[data-component-id='menu-toggle-button'])
            try {
                WebElement host = wait1.until(ExpectedConditions.presenceOfElementLocated(
                        By.cssSelector("ion-button[data-component-id='menu-toggle-button']")
                ));
                boolean hasShadow = hasShadowRoot(host, js);
                WebElement clickable = hasShadow
                        ? host.getShadowRoot().findElement(By.cssSelector("button.button-native, .button-inner"))
                        : firstChildOrSelf(host, "button.button-native, .button-inner");

                wait1.until(ExpectedConditions.invisibilityOfElementLocated(By.cssSelector("manh-overlay-container")));
                js.executeScript("arguments[0].click();", clickable);
                System.out.println("Menu toggle button clicked.");
            } catch (Exception e) {
                System.err.println("Error toggling menu: " + e.getMessage());
            }

            // 2) Type into the menu search input (ion-input[data-component-id='search-input'])
            try {
                WebElement inputHost = wait.until(ExpectedConditions.presenceOfElementLocated(
                        By.cssSelector("ion-input[data-component-id='search-input']")
                ));
                WebElement innerInput = findNativeInputInsideIonInputHost(inputHost, js);

                wait.until(ExpectedConditions.elementToBeClickable(innerInput));
                js.executeScript("arguments[0].scrollIntoView({block:'center'});", innerInput);
                innerInput.click();
                innerInput.sendKeys(Keys.chord(Keys.CONTROL, "a"));
                innerInput.sendKeys(Keys.DELETE);
                innerInput.clear();
                innerInput.sendKeys(MENU_ITEM_TEXT);
                System.out.println("✅ Menu Search typed: " + MENU_ITEM_TEXT);
            } catch (Exception e) {
                System.err.println("❌ Menu search input error: " + e.getMessage());
            }

            // 3) Click the menu item (native <button>) by ID or data-component-id or visible text
            boolean clicked = false;
            try {
                WebElement byId = wait.until(ExpectedConditions.presenceOfElementLocated(By.id(MENU_BUTTON_ID)));
                wait.until(ExpectedConditions.elementToBeClickable(byId));
                js.executeScript("arguments[0].scrollIntoView({block:'center'});", byId);
                js.executeScript("arguments[0].click();", byId);
                System.out.println("Clicked menu item by id: " + MENU_BUTTON_ID);
                clicked = true;
            } catch (Exception e) {
                System.err.println("Click by id failed: " + e.getMessage());
            }

            if (!clicked) {
                try {
                    WebElement byDataId = wait.until(ExpectedConditions.presenceOfElementLocated(
                            By.cssSelector("button[data-component-id='" + MENU_COMPONENT_ID + "']")
                    ));
                    wait.until(ExpectedConditions.elementToBeClickable(byDataId));
                    js.executeScript("arguments[0].scrollIntoView({block:'center'});", byDataId);
                    js.executeScript("arguments[0].click();", byDataId);
                    System.out.println("Clicked menu item by data-component-id: " + MENU_COMPONENT_ID);
                    clicked = true;
                } catch (Exception e) {
                    System.err.println("Click by data-component-id failed: " + e.getMessage());
                }
            }

            if (!clicked) {
                try {
                    WebElement byText = wait.until(ExpectedConditions.presenceOfElementLocated(
                            By.xpath("//button[.//ion-label[normalize-space(text())='" + MENU_ITEM_TEXT + "']]")
                    ));
                    wait.until(ExpectedConditions.elementToBeClickable(byText));
                    js.executeScript("arguments[0].scrollIntoView({block:'center'});", byText);
                    js.executeScript("arguments[0].click();", byText);
                    System.out.println("Clicked menu item by text: " + MENU_ITEM_TEXT);
                    clicked = true;
                } catch (Exception e) {
                    System.err.println("Click by text failed: " + e.getMessage());
                }
            }

            if (!clicked) {
                throw new RuntimeException("Failed to click the oLPNs 2.0 menu item.");
            }

            // 4) Try to close the menu if it stays open
            closeSideMenuIfOpen(driver);

            sleep(1000);
        }

        // ---------- One-shot per OLPN: Sort → Chevron → Input ----------
        private static void sortChevronAndEnterOlpnOnce(String olpn,WebDriver driver) {
            WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(time));
            JavascriptExecutor js = (JavascriptExecutor) driver;

            ensureFrameForOlpns(driver);

            // Guard overlays
            try { wait.until(ExpectedConditions.invisibilityOfElementLocated(By.cssSelector("manh-overlay-container"))); }
            catch (TimeoutException ignored) {}

            // A) Click the Filter/Sort button at index 3
            try {
                WebElement filterBtnHost = wait.until(ExpectedConditions.presenceOfElementLocated(
                        By.xpath("(//ion-button[contains(@class,'toggle-button')])[3]")
                ));
                clickIonButtonHost(filterBtnHost, js);
                System.out.println("Filter/Sort toggle clicked (index 3).");
            } catch (Exception e) {
                System.err.println("Filter/Sort button error: " + e.getMessage());
                e.printStackTrace(System.err);
            }

            try { sleep(400); } catch (InterruptedException ignored) {}

//Clear Filter
try {
  //  WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));

    WebElement clearAllBtn = wait.until(ExpectedConditions.elementToBeClickable(
            By.cssSelector("button[data-component-id='clear-all-btn']")
    ));

    clearAllBtn.click();
}
catch (Exception downErr) {
    // Fallback: try chevron-down icon
    System.err.println("Clear Filter " + downErr.getMessage());
}
//Expand Filter Again
            try {
                WebElement filterBtnHost = wait.until(ExpectedConditions.presenceOfElementLocated(
                        By.xpath("(//ion-button[contains(@class,'toggle-button')])[3]")
                ));
                clickIonButtonHost(filterBtnHost, js);
                System.out.println("Filter/Sort toggle clicked (index 3).");
            } catch (Exception e) {
                System.err.println("Filter/Sort button error: " + e.getMessage());
                e.printStackTrace(System.err);
            }




                // B) Chevron logic: if 'up' present, already expanded; else click 'down'
            try {
                List<WebElement> chevronUpButtons = driver.findElements(
                        By.cssSelector("ion-button[data-component-id='" + COMP_ID_CHEVRON_UP + "']")
                );
                if (!chevronUpButtons.isEmpty() && chevronUpButtons.get(0).isDisplayed()) {
                    System.out.println("Chevron-up present (expanded).");
                } else {
                    System.out.println("Chevron-up not present; expanding section.");

                    // Try explicit 'down' component-id first
                    try {
                        WebElement downHost = driver.findElement(By.cssSelector(
                                "ion-button[data-component-id='" + COMP_ID_CHEVRON_DOWN + "']"
                        ));
                        new WebDriverWait(driver, Duration.ofSeconds(time))
                                .until(ExpectedConditions.elementToBeClickable(downHost));
                        clickIonButtonHost(downHost, js);
                        System.out.println("Chevron-down button clicked.");
                    } catch (Exception downErr) {
                        // Fallback: try chevron-down icon
                        System.err.println("Chevron-down by component-id not found; trying icon fallback. Reason: " + downErr.getMessage());
                        try {
                            WebElement downIcon = new WebDriverWait(driver, Duration.ofSeconds(time))
                                    .until(ExpectedConditions.presenceOfElementLocated(
                                            By.cssSelector("ion-icon[name='chevron-down']")
                                    ));
                            js.executeScript("arguments[0].scrollIntoView({block:'center'});", downIcon);
                            try {
                                new WebDriverWait(driver, Duration.ofSeconds(time))
                                        .until(ExpectedConditions.elementToBeClickable(downIcon)).click();
                                System.out.println("Chevron-down icon clicked (host).");
                            } catch (Exception e2) {
                                boolean hasShadowIcon = hasShadowRoot(downIcon, js);
                                if (hasShadowIcon) {
                                    SearchContext sr = downIcon.getShadowRoot();
                                    WebElement innerSvg = sr.findElement(By.cssSelector(".icon-inner svg"));
                                    js.executeScript("arguments[0].click();", innerSvg);
                                    System.out.println("Chevron-down icon clicked (inner svg via JS).");
                                } else {
                                    js.executeScript("arguments[0].click();", downIcon);
                                    System.out.println("Chevron-down icon clicked via JS.");
                                }
                            }
                        } catch (Exception finalErr) {
                            System.err.println("Failed to expand via icon fallback: " + finalErr.getMessage());
                        }
                    }
                }
            } catch (Exception e) {
                System.err.println("Chevron logic error: " + e.getMessage());
                e.printStackTrace();
            }

            // C) Enter OLPN into input: ion-input[data-component-id='OlpnId']
            try {
                WebElement olpnHost = wait.until(ExpectedConditions.presenceOfElementLocated(
                        By.cssSelector("ion-input[data-component-id='" + COMP_ID_OLPN_INPUT + "']")
                ));
                js.executeScript("arguments[0].scrollIntoView({block:'center'});", olpnHost);

                WebElement input = findNativeInputInsideIonInputHost(olpnHost, js);
                wait.until(ExpectedConditions.elementToBeClickable(input));

                input.click();
                input.sendKeys(Keys.chord(Keys.CONTROL, "a"));
                input.sendKeys(Keys.DELETE);
                input.clear();
                input.sendKeys(olpn);
                input.sendKeys(Keys.ENTER);

                // Fire events for Ionic listeners
                js.executeScript("arguments[0].dispatchEvent(new InputEvent('input', {bubbles:true, composed:true}));", input);
                js.executeScript("arguments[0].dispatchEvent(new Event('change', {bubbles:true, composed:true}));", input);

                System.out.println("Entered OLPN: " + olpn);
            } catch (Exception e) {
                System.err.println("Failed to enter OLPN into filter: " + e.getMessage());
                e.printStackTrace();
            }
        }

        // ---------- Helpers ----------
        private static boolean hasShadowRoot(WebElement host, JavascriptExecutor js) {
            Object val = js.executeScript("return arguments[0] && arguments[0].shadowRoot != null;", host);
            return Boolean.TRUE.equals(val);
        }

        private static WebElement firstChildOrSelf(WebElement host, String css) {
            List<WebElement> kids = host.findElements(By.cssSelector(css));
            return kids.isEmpty() ? host : kids.get(0);
        }

        private static WebElement findNativeInputInsideIonInputHost(WebElement inputHost, JavascriptExecutor js) {
            if (hasShadowRoot(inputHost, js)) {
                SearchContext sr = inputHost.getShadowRoot();
                List<WebElement> candidates = sr.findElements(By.cssSelector("input.native-input, input"));
                if (!candidates.isEmpty()) return candidates.get(0);
            }
            // Light DOM fallback (matches your inspect: <label>…<div class='native-wrapper'><input class='native-input'>…)
            List<WebElement> lightCandidates = inputHost.findElements(
                    By.cssSelector("input.native-input, .native-wrapper input, label input, input")
            );
            if (!lightCandidates.isEmpty()) return lightCandidates.get(0);
            throw new NoSuchElementException("Native input not found inside ion-input host.");
        }

        private static void clickIonButtonHost(WebElement host, JavascriptExecutor js) {
            boolean hasShadow = hasShadowRoot(host, js);
            WebElement nativeBtn = hasShadow
                    ? host.getShadowRoot().findElement(By.cssSelector("button.button-native, .button-inner"))
                    : firstChildOrSelf(host, "button.button-native, .button-inner");

            js.executeScript("arguments[0].scrollIntoView({block:'center'});", nativeBtn);
            js.executeScript("arguments[0].click();", nativeBtn);
        }

        // Robust menu closer (clicks ion-icon inner SVG if shadow exists)
        private static void closeSideMenuIfOpen(WebDriver driver) {
            WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(time));
            JavascriptExecutor js = (JavascriptExecutor) driver;

            for (int attempt = 1; attempt <= 3; attempt++) {
                try {
                    System.err.println("Click by text failed: 11111111111");
                    try {
                        wait.until(ExpectedConditions.invisibilityOfElementLocated(By.cssSelector("manh-overlay-container")));
                    } catch (TimeoutException ignored) {
                        System.err.println("Click by text failed: 2222222222");
                    }



                                        WebElement closeHost = driver.findElement(By.id("close-menu-button"));
                    if (!closeHost.isDisplayed()) {
                        System.err.println("Click by text failed: 33333333");
                        sleep(250);
                        continue;
                    }



                    List<WebElement> closeButtons = driver.findElements(By.id("close-menu-button"));
                    System.out.println("GOT2");

                    if (!closeButtons.isEmpty() && closeButtons.get(0).isDisplayed()) {
                        System.out.println("GOT3");

                        // If present and visible, click it
                        closeButtons.get(0).click();
                        sleep(250);
//                        continue;
                    }


//                    WebElement closeHost = driver.findElement(By.id("close-menu-button"));
//                    if (!closeHost.isDisplayed()) {
//                        System.err.println("Click by text failed: 33333333");
//                        Thread.sleep(250);
//                        continue;
//                    }
//                    System.err.println("Click by text failed: 4444444444");
//                    js.executeScript("arguments[0].scrollIntoView({block:'center'});", closeHost);
//
//                    if (hasShadowRoot(closeHost, js)) {
//                        System.err.println("Click by text failed: 7777777777");
//                        SearchContext sr = closeHost.getShadowRoot();
//                        WebElement innerSvg = sr.findElement(By.cssSelector(".icon-inner svg"));
//                        try {
//                            wait.until(ExpectedConditions.elementToBeClickable(innerSvg)).click();
//                            System.err.println("Click by text failed: 55555555555");
//                        } catch (Exception e) {
//                            System.err.println("Click by text failed: 66666666666");
//                            js.executeScript("arguments[0].click();", innerSvg);
//                        }
//                    } else {
//                        try {
//                            System.err.println("Click by text failed: 888888888");
//                            wait.until(ExpectedConditions.elementToBeClickable(closeHost)).click();
//                        } catch (Exception e) {
//                            System.err.println("Click by text failed: 9999999999");
//                            js.executeScript("arguments[0].click();", closeHost);
//                        }
//                    }
//
//                    Thread.sleep(200);
//                    List<WebElement> stillThere = driver.findElements(By.id("close-menu-button"));
//                    if (stillThere.isEmpty() || !stillThere.get(0).isDisplayed()) {
//                        System.out.println("✅ Side menu closed.");
//                        return;
//                    } else {
//                        System.out.println("Close icon still visible; retrying (" + attempt + ").");
//                        Thread.sleep(250);
//                    }
//                } catch (NoSuchElementException e) {
//                    System.out.println("Side menu close icon not found — likely already closed.");
//                    return;
//                } catch (InterruptedException ignored) {
                } catch (Exception e) {
                    System.err.println("Close menu attempt " + attempt + " failed: " + e.getMessage());
                }

                    System.out.println("⚠ Attempted to close menu 3 times; icon still visible.");

            }
        }

        // ---------- Frame handling ----------
        private static final By[] OLPN_MARKERS = new By[] {
                By.xpath("(//ion-button[contains(@class,'toggle-button')])[3]"),
                By.cssSelector("ion-button[data-component-id='" + COMP_ID_CHEVRON_UP + "']"),
                By.cssSelector("ion-input[data-component-id='" + COMP_ID_OLPN_INPUT + "']")
        };

        private static boolean isMarkerPresent(WebDriver driver,By... markers) {
            for (By m : markers) {
                if (!driver.findElements(m).isEmpty()) return true;
            }
            return false;
        }

        private static boolean switchToFrameRecursive(WebDriver driver,By... markers) {
            if (isMarkerPresent(driver,markers)) return true;

            List<WebElement> frames = driver.findElements(By.cssSelector("iframe, frame"));
            for (WebElement frame : frames) {
                try {
                    driver.switchTo().frame(frame);
                    new WebDriverWait(driver, SHORT_WAIT).until(d -> true); // small yield
                    if (switchToFrameRecursive(driver,markers)) return true;
                    driver.switchTo().parentFrame();
                } catch (Exception ignored) {
                    driver.switchTo().parentFrame();
                }
            }
            return false;
        }

        private static void ensureFrameForOlpns(WebDriver driver) {
            try {
                driver.switchTo().defaultContent();
                boolean found = switchToFrameRecursive(driver,OLPN_MARKERS);
                if (found) {
                    System.out.println("✅ Switched into frame containing OLPNs controls.");
                } else {
                    driver.switchTo().defaultContent();
                    System.out.println("ℹ OLPN markers not found; staying in default content.");
                }
            } catch (Exception e) {
                System.out.println("ensureFrameForOlpns() error: " + e.getMessage());
                driver.switchTo().defaultContent();
            }
        }

        // ---------- Excel reading ----------
        public static String[] fetchOlpnsForDeclaredTestcase(String EXCEL_PATH,String TESTCASE_VALUE) {
            System.out.println("No OLPNs found for testcase: " + EXCEL_PATH);
            final String sheetName = "Tasks";
            final Set<String> olpns = new LinkedHashSet<>(); // preserve order & dedupe
            final DataFormatter formatter = new DataFormatter();

            try (FileInputStream fis = new FileInputStream(EXCEL_PATH);
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
                Integer olpnsCol = headerIndex.get("OLPNs");

                if (tcCol == null) throw new IllegalStateException("Column 'Testcase' not found.");
                if (olpnsCol == null) throw new IllegalStateException("Column 'OLPNs' not found.");

                int firstDataRow = headerRowIdx + 1;
                for (int r = firstDataRow; r <= sheet.getLastRowNum(); r++) {
                    Row row = sheet.getRow(r);
                    if (row == null) continue;

                    String tc = cleanupInvisible(safeFormat(formatter, row.getCell(tcCol)));
                    if (!TESTCASE_VALUE.equals(tc)) continue;

                    String raw = cleanupInvisible(safeFormat(formatter, row.getCell(olpnsCol)));
                    if (raw.isEmpty()) continue;

                    // Support multi-values: comma, semicolon, whitespace
                    String[] parts = raw.split("[,;\\s]+");
                    for (String p : parts) {
                        String v = cleanupInvisible(p);
                        if (!v.isEmpty()) {
                            olpns.add(v);
                        }
                    }
                }

                if (olpns.isEmpty()) {
                    System.out.println("No OLPNs found for Testcase: " + TESTCASE_VALUE + " in sheet '" + sheetName + "'");
                } else {
                    System.out.println("OLPNs for Testcase " + TESTCASE_VALUE + ":");
                    int i = 1;
                    for (String v : olpns) {
                        System.out.println((i++) + ". " + v);
                    }
                }
            } catch (Exception e) {
                throw new RuntimeException("Failed to read Excel and fetch OLPNs: " + e.getMessage(), e);
            }

            return olpns.toArray(new String[0]);
        }

        private static String safeFormat(DataFormatter formatter, Cell cell) {
            if (cell == null) return "";
            return formatter.formatCellValue(cell).trim();
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

    public static void selectCardAndOpenTask(WebDriver driver) {
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(time));
        JavascriptExecutor js = (JavascriptExecutor) driver;

        // 1) Select the first Card View row (primary, focusable)
        try {
            By cardViewBy = By.cssSelector("card-view[data-component-id='Card-View'] .card-row.primary[tabindex='0']");
            WebElement cardView = wait.until(ExpectedConditions.elementToBeClickable(cardViewBy));
            js.executeScript("arguments[0].scrollIntoView({block:'center'});", cardView);
            cardView.click();
            System.out.println("✅ Card view selected.");
        } catch (StaleElementReferenceException staleEx) {
            System.out.println("⚠️ Stale element detected. Retrying Card View selection...");
            try {
                By cardViewBy = By.cssSelector("card-view[data-component-id='Card-View'] .card-row.primary[tabindex='0']");
                WebElement cardViewRetry = wait.until(ExpectedConditions.elementToBeClickable(cardViewBy));
                ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView({block:'center'});", cardViewRetry);
                cardViewRetry.click();
                System.out.println("✅ Card view selected after retry.");
            } catch (Exception retryEx) {
                System.out.println("❌ Retry failed for Card View: " + retryEx.getMessage());
            }
        } catch (Exception e) {
            System.out.println("❌ Failed to select the Card View: " + e.getMessage());
        }

        // 2) Click "Related Links" button
        try {
            WebDriverWait waitShort = new WebDriverWait(driver, Duration.ofSeconds(time));
            By relatedLinksButtonLocator = By.cssSelector("button[data-component-id='relatedLinks']");
            WebElement relatedLinksButton = waitShort.until(ExpectedConditions.elementToBeClickable(relatedLinksButtonLocator));
            ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView({block:'center'});", relatedLinksButton);
            ((JavascriptExecutor) driver).executeScript("arguments[0].click();", relatedLinksButton);
            System.out.println("✅ Related Links button clicked.");
        } catch (Exception e) {
            System.out.println("❌ Failed to click Related Links: " + e.getMessage());
        }


        try {
            WebDriverWait waitShort = new WebDriverWait(driver, Duration.ofSeconds(60));
            By taskLinkBy = By.xpath("//a[normalize-space(text())='Task']");
            WebElement taskLink = waitShort.until(ExpectedConditions.elementToBeClickable(taskLinkBy));
            ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView({block:'center'});", taskLink);
            ((JavascriptExecutor) driver).executeScript("arguments[0].click();", taskLink);
            System.out.println("✅ Task link clicked.");
            sleep(10000);
            captureScreenshot("Olpns-Tasks",driver);
            captureAllCardsScreenshots(driver);
            DocPathManager.saveSharedDocument();
            sleep(5000);
            navigateTillWaveRuns1(driver);

        } catch (TimeoutException te) {
            System.out.println("ℹ️ Task link not found within timeout.");
        } catch (Exception e) {
            System.out.println("❌ Error clicking Task link: " + e.getMessage());
        }






    }


    public static void navigateTillWaveRuns1(WebDriver driver) throws InterruptedException, IOException {
        WebElement waveRunsLink = driver.findElement(By.cssSelector("a[title='oLPNs 2.0']"));
        ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", waveRunsLink);
        waveRunsLink.click();
        sleep(5000);
        System.out.println("oLPNs 2.0 clicked.");
    }




}

