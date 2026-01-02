
package MA_MSG_Suite_OB;
import com.google.gson.*;
import io.github.bonigarcia.wdm.WebDriverManager;
import okhttp3.*;
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

import java.io.*;
import java.nio.file.*;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.*;

import static javax.swing.UIManager.getString;

public class Main6a_WaveToReleaseTask {
    public static String statusText;
    // ❌ Removed static XWPFDocument to avoid carry-over across runs
    public static WebDriver driver;
    public static String waveNumber;
    // We’ll avoid using a static docPath and pass it around where needed
    public static Map<String, String> opsToWaveMap = new HashMap<>(); // OPS -> WaveNum
    public static Map<String, List<String>> waveTaskMap = new HashMap<>(); // Wave -> Task IDs
    public static Map<String, List<String>> waveOlpnMap = new HashMap<>(); // Wave -> OLPN IDs
    public static int time = 60;



    // =========================
    // MAIN (updated & complete)
    // =========================
    public static void main(String filePath, String testcase, String env) {
        // Fresh document per run
        XWPFDocument document = new XWPFDocument();
        String docPathLocal = DocPathManager.getOrCreateDocPath(filePath, testcase);

        // Reset static state that may leak across IDE runs
        waveNumber = null;
        opsToWaveMap.clear();
        waveTaskMap.clear();
        waveOlpnMap.clear();

        try {
            // Unique, timestamped output path (same directory as Excel)
           // docPathLocal = buildDocPath(filePath, testcase);
          // String docPathLocal = DocPathManager.getOrCreateDocPath(filePath, testcase);
            System.out.println("Output doc: " + docPathLocal);

            // Extract OPS
            String path = filePath;
            String testcaseToRun = testcase;
            String ops = getOpsForTestcase(path, testcaseToRun);
            System.out.println("Processing OPS: " + ops);

            // WebDriver & Login
            WebDriverManager.chromedriver().setup();
            ChromeOptions options = new ChromeOptions();
            options.addArguments("--start-maximized");
            driver = new ChromeDriver(options);
            driver.manage().window().maximize();
            Main1_URL_Login1 login1 = new Main1_URL_Login1(driver, env);
            login1.execute();
            System.out.println("login done:\n");

            // Add a run header to the document
            addRunHeader(document, testcaseToRun, null);

            // Run OPS -> generate wave
            try {
                RunOPS(ops); // sets global waveNumber
                WavestatusWait(filePath, testcase, document, docPathLocal); // status + capture + save
            } catch (Exception e) {
                System.err.println("Error while processing OPS: " + ops);
                e.printStackTrace();
            }

            // Generate Task & downstream flow
            GenerateTask();
            System.out.println("Executing WaveToExcel after 20 sec");
            Thread.sleep(20000);
            Main6b_WaveToExcel.main(waveNumber, path, testcaseToRun);
            Thread.sleep(2000);

            WaveSelectionAndRelatedlinks();
            Thread.sleep(2000);
            Allocation(document, docPathLocal);
            Thread.sleep(2000);

            navigateTillWaveRuns1();
            Thread.sleep(2000);
            WaveSelectionAndRelatedlinks();
            Thread.sleep(2000);
            clickOLPNs(document, docPathLocal);
            Thread.sleep(2000);

            navigateTillWaveRuns1();
            Thread.sleep(2000);
            WaveSelectionAndRelatedlinks();
            Thread.sleep(2000);
            clickTasks(document, docPathLocal);
            Thread.sleep(2000);
            ReleaseTask(document, docPathLocal);
          DocPathManager.saveSharedDocument();
            if (driver != null) {
                driver.quit();
            }
            System.out.println(" Waving Done for Testcase"+testcase);
        } catch (Exception e) {
            System.err.println("❌ Error occurred: " + e.getMessage());
            e.printStackTrace();
            updateOpsStatus(filePath, testcase, "Failed");
            // Best-effort save on failure
            if (docPathLocal != null) {
                DocPathManager.saveSharedDocument();
            }
        } finally {
            try {
                document.close();
            } catch (IOException ignore) {
            }
            if (driver != null) {
                try {
                    driver.quit();
                } catch (Exception ignore) {
                }
            }
        }
    }

    // =========================
    // Helpers (new/updated)
    // =========================

    // Build unique doc path (timestamped) in same dir as Excel
    public static String buildDocPath(String excelPathStr, String baseName) {
        Path excelPath = Paths.get(excelPathStr);
        Path parent = excelPath.getParent() != null ? excelPath.getParent() : Paths.get(".");
        String stamp = new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date());
        String unique = baseName + "_" + stamp + ".docx";
        return parent.resolve(unique).toString();
    }

    // Add a header for readability
    public static void addRunHeader(XWPFDocument document, String testcase, String waveNum) {
        XWPFParagraph p = document.createParagraph();
        XWPFRun r = p.createRun();
        r.setBold(true);
        r.setFontSize(12);
        r.setText("Run Summary | Testcase: " + (testcase != null ? testcase : "N/A")
                + " | Wave: " + (waveNum != null ? waveNum : "N/A"));
        r.addBreak();
    }




//    public static void captureScreenshot(String fileName) {
//        try {
//            File srcFile = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
//            try (FileInputStream fis = new FileInputStream(srcFile)) {
//                XWPFDocument document = DocPathManager.getSharedDocument();
//                XWPFParagraph paragraph = document.createParagraph();
//                XWPFRun run = paragraph.createRun();
//                run.setText("Screenshot: " + fileName);
//                run.addBreak();
//                run.addPicture(fis,
//                        Document.PICTURE_TYPE_PNG,
//                        fileName + ".png",
//                        Units.toEMU(500),
//                        Units.toEMU(300));
//            }
//            System.out.println("Screenshot added to document: " + fileName);
//        } catch (Exception e) {
//            System.out.println("Error capturing screenshot: " + e.getMessage());
//        }
//    }
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






    // =========================
    // Existing methods (adapted to pass document & docPath where needed)
    // =========================

    public static String getOpsForTestcase(String path, String testcaseToRun) throws IOException {
        try (FileInputStream fis = new FileInputStream(path);
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheet("OPS_Tab");
            if (sheet == null) {
                System.out.println("❌ Sheet 'OPS & Wave Num' not found!");
                return null;
            }
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;
                String testcase = getCellValueAsString(row.getCell(0));
                if (testcase != null && testcase.trim().equalsIgnoreCase(testcaseToRun)) {
                    return getCellValueAsString(row.getCell(1)); // Column 1 = OPS
                }
            }
        }
        return null; // No matching testcase found
    }

    private static String getCellValueAsString(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                return String.valueOf((int) cell.getNumericCellValue()).trim();
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue()).trim();
            case FORMULA:
                return cell.getCellFormula().trim();
            default:
                return "";
        }
    }

    public static void fetchAndWriteTaskOlpnData(String token, String filePath) throws IOException {
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet taskSheet = workbook.getSheet("Tasks");
            if (taskSheet == null) {
                System.err.println("❌ Sheet 'Tasks' not found.");
                return;
            }
            Row row1 = taskSheet.getRow(1);
            if (row1 != null) {
                System.out.print("Row 1 values: ");
                for (int c = 0; c < row1.getLastCellNum(); c++) {
                    Cell cell = row1.getCell(c);
                    String value = (cell == null) ? "" : cell.toString();
                    System.out.print(value + " \n ");
                }
                System.out.println();
            }
            int taskRowIndex = 1;
            for (int i = 1; i <= taskSheet.getLastRowNum(); i++) {
                Row row = taskSheet.getRow(i);
                if (row == null) continue;
                Cell waveCell = row.getCell(0);
                if (waveCell == null) continue;
                String waveNumber = waveCell.getStringCellValue().trim();
                if (waveNumber.isEmpty()) continue;

                List<String> taskIds = fetchIds(waveNumber, "task", "GenerationNumberId", token, "com-manh-cp-task", "TaskId");
                List<String> olpnIds = fetchIds(waveNumber, "olpn", "OrderPlanningRunId", token, "com-manh-cp-pickpack", "OlpnId");

                System.out.println("Wave: " + waveNumber + " \n Tasks: " + taskIds + " \n OLPNs: " + olpnIds);
                waveTaskMap.put(waveNumber, taskIds);
                waveOlpnMap.put(waveNumber, olpnIds);

                String olpnId = olpnIds.isEmpty() ? "" : olpnIds.get(0);
                for (String taskId : taskIds) {
                    Row newRow = taskSheet.getRow(taskRowIndex);
                    if (newRow == null) {
                        newRow = taskSheet.createRow(taskRowIndex);
                    }
                    newRow.createCell(1).setCellValue(waveNumber); // Column B
                    newRow.createCell(2).setCellValue(taskId);     // Column C
                    newRow.createCell(8).setCellValue(olpnId);     // Column I
                    newRow.createCell(3).setCellValue("");         // Column D
                    taskRowIndex++;
                }
            }
            try (FileOutputStream fos = new FileOutputStream(filePath)) {
                workbook.write(fos);
            }
            System.out.println("✅ Task IDs and OLPN IDs written starting from row 2.");
        }
    }

    public static List<String> fetchIds(String waveNumber, String viewName, String attributeId, String token, String componentName, String idKey) {
        List<String> ids = new ArrayList<>();
        int maxRetries = 3;
        for (int attempt = 1; attempt <= maxRetries; attempt++) {
            try {
                OkHttpClient client = new OkHttpClient();
                MediaType mediaType = MediaType.parse("application/json");

                JsonObject filter = new JsonObject();
                filter.addProperty("ViewName", viewName);
                filter.addProperty("AttributeId", attributeId);
                filter.add("FilterValues", new Gson().toJsonTree(List.of(waveNumber.trim())));
                filter.addProperty("requiredFilter", false);
                filter.addProperty("Operator", "=");

                JsonObject body = new JsonObject();
                body.addProperty("ViewName", viewName.equalsIgnoreCase("olpn") ? "DMOlpn" : "Task");
                body.add("Filters", new Gson().toJsonTree(List.of(filter)));
                body.addProperty("ComponentName", componentName);
                body.addProperty("Size", 100);
                body.addProperty("TimeZone", "Europe/Paris");

                RequestBody requestBody = RequestBody.create(mediaType, body.toString());
                Request request = new Request.Builder()
                        .url("https://ujdss.sce.manh.com/dmui-facade/api/dmui-facade/entity/search")
                        .post(requestBody)
                        .addHeader("Content-Type", "application/json")
                        .addHeader("Authorization", "Bearer " + token)
                        .addHeader("SelectedOrganization", "HEERLEN51")
                        .addHeader("SelectedLocation", "HEERLEN51")
                        .build();

                Response response = client.newCall(request).execute();
                String responseBody = response.body() != null ? response.body().string() : "No response body";

                System.out.println("\n🔍 Attempt " + attempt + " for wave: " + waveNumber);
                System.out.println("Request Body: " + body.toString());
                System.out.println("Response Code: " + response.code());
                System.out.println("Response Body: " + responseBody);

                if (!response.isSuccessful()) {
                    System.err.println("❌ Failed to fetch " + idKey + " (HTTP " + response.code() + ")");
                    if (response.code() == 500 && attempt < maxRetries) {
                        System.out.println("🔁 Retrying...");
                        Thread.sleep(1000);
                        continue;
                    } else {
                        break;
                    }
                }

                JsonObject json = JsonParser.parseString(responseBody).getAsJsonObject();
                if (json.has("data")) {
                    JsonObject dataObject = json.getAsJsonObject("data");
                    if (dataObject.has("Results")) {
                        JsonArray results = dataObject.getAsJsonArray("Results");
                        for (JsonElement element : results) {
                            JsonObject obj = element.getAsJsonObject();
                            if (obj.has(idKey)) {
                                ids.add(obj.get(idKey).getAsString());
                            }
                        }
                    }
                }
                break;
            } catch (Exception e) {
                System.err.println("❌ Exception fetching " + idKey + " (Attempt " + attempt + "): " + e.getMessage());
            }
        }
        System.out.println("✅ Fetched " + ids.size() + " " + idKey + "(s) for wave: " + waveNumber);
        return ids;
    }

    public static void SearchMenu(String Keyword, String id) {
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
    }

    public static void RunOPS(String OPS) throws InterruptedException, IOException {
        System.out.println("Run OPS starts");
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(time));
        Thread.sleep(5000);
        System.out.println("OPS: " + OPS);
        SearchMenu("Order Planning Strategy", "orderPlanningStrategy");

//        WebDriverWait wait1 = new WebDriverWait(driver, Duration.ofSeconds(time));
//        System.out.println("closing Menu while RunOPS");
//        WebElement closeIcon = wait.until(ExpectedConditions.visibilityOfElementLocated(
//                By.cssSelector("ion-icon[data-component-id='close']")
//        ));
//        ((JavascriptExecutor) driver).executeScript("arguments[0].click();", closeIcon);
//        System.out.println("closing Menu while RunOPS Done");


        try {

// Try to find the close icon without throwing an exception

            List<WebElement> closeIcons = driver.findElements(
                    By.cssSelector("ion-icon[data-component-id='close']")
            );
            if (!closeIcons.isEmpty() && closeIcons.get(0).isDisplayed()) {
                WebElement closeIcon = closeIcons.get(0);

// Use JS click for Ionic components

                ((JavascriptExecutor) driver).executeScript("arguments[0].click();", closeIcon);
                System.out.println("Closing menu while RunOPS done");
            } else {
                System.out.println("Menu already closed");
            }
        } catch (Exception e) {
            System.out.println("Error while checking/closing menu: " + e.getMessage());
        }






        System.out.println("Order Planning Strategy button clicked.");
        Thread.sleep(5000);
        try {
            WebElement filterBtnHost = wait.until(ExpectedConditions.elementToBeClickable(
                    By.xpath("(//ion-button[contains(@class,\"toggle-button\")])[3]")
            ));
            JavascriptExecutor jse = (JavascriptExecutor) driver;
            WebElement expandButton = (WebElement) jse.executeScript(
                    "let btn = arguments[0].shadowRoot.querySelector('.button-inner');" +
                            " btn.click(); ", filterBtnHost);

            System.out.println("Order Planning Strategy click on Filter button.");
        } catch (Exception e) {
            System.err.println("Error: " + e.getMessage());
            e.printStackTrace(System.err);
        }
        Thread.sleep(5000);
        wait = new WebDriverWait(driver, Duration.ofSeconds(time));
        List<WebElement> chevronUpButtons = driver.findElements(
                By.cssSelector("ion-button[data-component-id='OrderPlanningStrategy-Name-chevron-up']")
        );
        System.out.println("GOT1");

        if (!chevronUpButtons.isEmpty()) {
            System.out.println("Chevron-up is already present. Skipping click.");
        } else {
            // Try to find the close-menu-button
            List<WebElement> closeButtons = driver.findElements(By.id("close-menu-button"));
            System.out.println("GOT2");

            if (!closeButtons.isEmpty() && closeButtons.get(0).isDisplayed()) {
                System.out.println("GOT3");

                // If present and visible, click it
                closeButtons.get(0).click();
                try {
                    // Otherwise, wait for expandButton1 and click it


                    WebElement expandButton1 = driver.findElement(
                            By.cssSelector("ion-button[data-component-id='OrderPlanningStrategy-Name-chevron-down']")
                    );

// Wait until clickable
                    wait.until(ExpectedConditions.elementToBeClickable(expandButton1));

// Now click


                    System.out.println("GOT5");
                    Thread.sleep(3000);
                    expandButton1.click();
                    System.out.println("Chevron-down button clicked using native click.");


                } catch (Exception e) {
                    System.err.println("Error: " + e.getMessage());
                    e.printStackTrace(System.err);
                }


            } else {

                try {
                    // Otherwise, wait for expandButton1 and click it


                    WebElement expandButton1 = driver.findElement(
                            By.cssSelector("ion-button[data-component-id='OrderPlanningStrategy-Name-chevron-down']")
                    );


                    wait.until(ExpectedConditions.elementToBeClickable(expandButton1));



                    System.out.println("GOT4");
                    Thread.sleep(3000);
                    expandButton1.click();
                    System.out.println("Chevron-down button clicked using native click.");


                } catch (Exception e) {
                    System.err.println("Error: " + e.getMessage());
                    e.printStackTrace(System.err);
                }
            }
        }


        try {
            Thread.sleep(8000);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }


        try {
//        WebElement planningStrategyInputField = wait.until(
//                ExpectedConditions.elementToBeClickable(By.xpath(
//                        "//ion-input[@data-component-id='PlanningStrategyId']//input"
//                ))
//        );
            // Locate the native input inside ion-input using data-component-id
            WebElement planningStrategyInputField = wait.until(
                    ExpectedConditions.visibilityOfElementLocated(
                            By.cssSelector("ion-input[data-component-id='PlanningStrategyId'] input")
                    )
            );







            if (OPS != null && !OPS.isEmpty()) {
                planningStrategyInputField.click();
                planningStrategyInputField.sendKeys(Keys.CONTROL + "a");
                planningStrategyInputField.sendKeys(Keys.DELETE);
                try {
                    Thread.sleep(3000);
                } catch (InterruptedException e) {
                    e.printStackTrace();
                }
                planningStrategyInputField.sendKeys(Keys.CONTROL + "a");
                planningStrategyInputField.sendKeys(Keys.DELETE);
                planningStrategyInputField.sendKeys(OPS);
                planningStrategyInputField.sendKeys(Keys.ENTER);
            }
            System.out.println("Planning Strategy ID entered: " + OPS);
        } catch (Exception e) {
            System.err.println("Error: " + e.getMessage());
            e.printStackTrace(System.err);
        }
        try {
            Thread.sleep(3000);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }

        WebElement actionClosedButton = driver.findElement(
                By.cssSelector("button[data-component-id='action-closed']")
        );
        JavascriptExecutor js3 = (JavascriptExecutor) driver;
        js3.executeScript("arguments[0].scrollIntoView(true);", actionClosedButton);
        try {
            Thread.sleep(500);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
        js3.executeScript("arguments[0].click();", actionClosedButton);
        System.out.println("Action Closed button clicked.");

        try {
            Thread.sleep(3000);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }

        WebElement runWaveButton = driver.findElement(
                By.cssSelector("button[data-component-id='RunWave']")
        );
        JavascriptExecutor js4 = (JavascriptExecutor) driver;
        js4.executeScript("arguments[0].scrollIntoView(true);", runWaveButton);
        try {
            Thread.sleep(500);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
        js4.executeScript("arguments[0].click();", runWaveButton);
        System.out.println("Run Wave button clicked.");

        WebDriverWait wait5 = new WebDriverWait(driver, Duration.ofSeconds(time));
        WebElement toastMessage = wait5.until(ExpectedConditions.visibilityOfElementLocated(
                By.cssSelector(".toast-message .text-wrap")
        ));
        String popupText = toastMessage.getText();
        waveNumber = popupText.replaceAll(".*Wave\\s+(\\S+)\\s+is submitted.*", "$1");
        System.out.println("Wave Number: " + waveNumber);
        opsToWaveMap.put(OPS, waveNumber);
    }

    public static void WavestatusWait(String filePath, String testcase, XWPFDocument document, String docPathLocal)
            throws InterruptedException, IOException {
        System.out.println("WaveStatus starts");
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(time));
        JavascriptExecutor js = (JavascriptExecutor) driver;

        Thread.sleep(3000);
        SearchMenu("Wave Run", "OrderPlanningRunStrategy");


        Thread.sleep(5000);

        try {
            WebElement filterBtnHost = wait.until(ExpectedConditions.elementToBeClickable(
                    By.xpath("(//ion-button[contains(@class,\"toggle-button\")])[3]")
            ));
            JavascriptExecutor jse = (JavascriptExecutor) driver;
            WebElement expandButton = (WebElement) jse.executeScript(
                    "let btn = arguments[0].shadowRoot.querySelector('.button-inner');" +
                            " btn.click(); ", filterBtnHost);
            System.out.println("Clicked toggle button");

        }
        catch (Exception e) {
            System.err.println("Error: " + e.getMessage());
            e.printStackTrace(System.err);
        }


//
//        WebElement filterBtnHost = wait.until(ExpectedConditions.elementToBeClickable(
//                By.xpath("(//ion-button[contains(@class,\"toggle-button\")])[3]")
//        ));
//        JavascriptExecutor jse = (JavascriptExecutor) driver;
//        WebElement expandButton = (WebElement) jse.executeScript(
//                "let btn = arguments[0].shadowRoot.querySelector('.button-inner');" +
//                        " btn.click(); ", filterBtnHost);

        Thread.sleep(5000);

        try {
            WebElement expandButton1 = wait.until(
                    ExpectedConditions.elementToBeClickable(
                            By.cssSelector("ion-button[data-component-id='OrderPlanningRunStrategy-Waverun-chevron-down']")
                    )
            );
            ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", expandButton1);

            System.out.println("Wave Run chevron-down button clicked in 10 sec");

            Thread.sleep(5000);
            expandButton1.click();
            System.out.println("Wave Run chevron-down button clicked using native click.");

        }
        catch (Exception e) {
            System.err.println("Error: " + e.getMessage());
            e.printStackTrace(System.err);
        }


        Thread.sleep(3000);


        try{
            WebElement orderPlanningRunInputField = wait.until(
                    ExpectedConditions.elementToBeClickable(By.xpath("//ion-input[@data-component-id='OrderPlanningRunId-lookup-dialog-filter-input']//input"))
            );
            if (waveNumber != null && !waveNumber.isEmpty()) {
                orderPlanningRunInputField.click();
                orderPlanningRunInputField.clear();
                orderPlanningRunInputField.sendKeys(waveNumber);
                orderPlanningRunInputField.sendKeys(Keys.ENTER);
            }
            System.out.println("Order Planning Run ID entered: " + waveNumber);



        }
        catch (Exception e) {
            System.err.println("Error: " + e.getMessage());
            e.printStackTrace(System.err);
        }






        int maxRetries = 40;
        for (int attempt = 1; attempt <= maxRetries; attempt++) {
            System.out.println("⏳ Waiting 30 seconds before refreshing...");
            try {
                Thread.sleep(10000);

                if(attempt == 1)
                {
                    DocPathManager.captureScreenshot("Wave Status",driver);
                    //  captureAllCardsScreenshots(document);
                     DocPathManager.saveSharedDocument();
                    System.out.println("🔁 Status is Started screenshot done");
                }else{
                    System.out.println("🔁 Status is Started screenshot already taken");
                }
                Thread.sleep(20000);


            } catch (InterruptedException e) {
                throw new RuntimeException(e);
            }

            WebElement refreshHost = driver.findElement(By.cssSelector("ion-button.refresh-button"));
            js = (JavascriptExecutor) driver;
            WebElement refreshButton = (WebElement) js.executeScript(
                    "return arguments[0].shadowRoot.querySelector('button.button-native')", refreshHost);
            refreshButton.click();
            System.out.println("🔄 Refresh button clicked.");

            Thread.sleep(3000);

            WebElement statusElement = wait.until(ExpectedConditions.visibilityOfElementLocated(
                    By.cssSelector("div[data-component-id='PlanningStatusDescription']")
            ));
            statusText = statusElement.getText().trim();
            System.out.println("📌 Planning Status: " + statusText);

            if ("Completed".equalsIgnoreCase(statusText)) {
                System.out.println("✅ Status is Completed: OK " + statusText);
                updateOpsStatus(filePath, testcase, statusText);
                break;
            } else if ("Cancelled".equalsIgnoreCase(statusText)) {
                System.out.println("❌ Status is Cancelled: skipping further actions.");
                updateOpsStatus(filePath, testcase, statusText);
                Thread.sleep(3000);
                DocPathManager.captureScreenshot("Wave Status",driver);
              //  captureAllCardsScreenshots(document);
                 DocPathManager.saveSharedDocument();
                return;
            } else if ("Started".equalsIgnoreCase(statusText) && attempt < maxRetries) {
                System.out.println("🔁 Status is Started: will retry after another minute...");

                updateOpsStatus(filePath, testcase, statusText);
            } else {
                System.out.println("⚠️ Status is still Started after retry or unknown status.");
                updateOpsStatus(filePath, testcase, statusText);
                return;
            }
        }
        Thread.sleep(3000);
        DocPathManager.captureScreenshot("Wave Status",driver);
      //  captureAllCardsScreenshots(document);
         DocPathManager.saveSharedDocument();
    }

    public static void WaveSelectionAndRelatedlinks() throws InterruptedException, IOException {
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(time));
        JavascriptExecutor js = (JavascriptExecutor) driver;
        try {
            WebElement cardView = wait.until(ExpectedConditions.elementToBeClickable(
                    By.cssSelector("card-view[data-component-id='Card-View'] .card-row.primary[tabindex='0']")
            ));
            js.executeScript("arguments[0].scrollIntoView(true);", cardView);
            cardView.click();
            System.out.println("Card view selected.");
        } catch (StaleElementReferenceException staleEx) {
            System.out.println("Stale element detected. Retrying...");
            WebElement cardViewRetry = wait.until(ExpectedConditions.elementToBeClickable(
                    By.cssSelector("card-view[data-component-id='Card-View'] .card-row.primary[tabindex='0']")
            ));
            js.executeScript("arguments[0].scrollIntoView(true);", cardViewRetry);
            cardViewRetry.click();
            System.out.println("Card view selected after retry.");
        } catch (Exception e) {
            System.out.println("Failed to select the card view: " + e.getMessage());
        }

        WebDriverWait waitShort = new WebDriverWait(driver, Duration.ofSeconds(time));
        By relatedLinksButtonLocator = By.cssSelector("button[data-component-id='relatedLinks']");
        WebElement relatedLinksButton = waitShort.until(ExpectedConditions.elementToBeClickable(relatedLinksButtonLocator));
        js = (JavascriptExecutor) driver;
        js.executeScript("arguments[0].click();", relatedLinksButton);
        System.out.println("Related Links button clicked.");
    }

    public static void Allocation(XWPFDocument document, String docPathLocal) throws InterruptedException, IOException {
        System.out.println("Allocation start");
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(time));
        JavascriptExecutor js = (JavascriptExecutor) driver;
        By allocationsLocator = By.xpath("//ion-item[@data-component-id='Allocations']//a[text()='Allocations']");
        WebElement allocationsLink = wait.until(ExpectedConditions.elementToBeClickable(allocationsLocator));
        js.executeScript("arguments[0].click();", allocationsLink);
        System.out.println("Allocations clicked.");
        Thread.sleep(5000);

        // wait = new WebDriverWait(driver, Duration.ofSeconds(20));
        WebElement cardPanel = wait.until(ExpectedConditions.visibilityOfElementLocated(
                By.xpath("/html/body/app-root/ion-app/div/ion-split-pane/ion-router-outlet/screen-page/div/div/div[2]/div/ion-content/card-panel")
        ));
        System.out.println("✅ card-panel is visible. Proceeding with next actions...");
        try {
            Thread.sleep(8000);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }

        DocPathManager.captureScreenshot("Allocations",driver);
        DocPathManager.captureAllCardsScreenshots(driver);
         DocPathManager.saveSharedDocument();
    }

    public static void navigateTillWaveRuns() throws InterruptedException, IOException {
        try {
            Thread.sleep(3000);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
        SearchMenu("Wave Run", "OrderPlanningRunStrategy");

        JavascriptExecutor js = (JavascriptExecutor) driver;
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(time));

        WebElement waveRunsButton = wait.until(ExpectedConditions.elementToBeClickable(By.id("OrderPlanningRunStrategy")));
        JavascriptExecutor js2 = (JavascriptExecutor) driver;
        js2.executeScript("arguments[0].click();", waveRunsButton);
        System.out.println("✅ Click on Wave Run");

        try {
            Thread.sleep(3000);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }

        try {
            //  wait = new WebDriverWait(driver, Duration.ofSeconds(10));
            WebElement closeButton = wait.until(ExpectedConditions.elementToBeClickable(By.id("close-menu-button")));
            closeButton.click();
            System.out.println("✅ Close Button");
        } catch (TimeoutException e) {
            System.out.println("Close menu button not found or already closed.");
        }

        Thread.sleep(5000);

        WebElement filterBtnHost = wait.until(ExpectedConditions.elementToBeClickable(
                By.xpath("(//ion-button[contains(@class,\"toggle-button\")])[3]")
        ));
        JavascriptExecutor jse = (JavascriptExecutor) driver;
        WebElement expandButton = (WebElement) jse.executeScript(
                "let btn = arguments[0].shadowRoot.querySelector('.button-inner');" +
                        " btn.click(); ", filterBtnHost);

        try {
            Thread.sleep(3000);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
        Thread.sleep(5000);


        WebElement expandButton1 = wait.until(
                ExpectedConditions.elementToBeClickable(
                        By.cssSelector("ion-button[data-component-id='OrderPlanningRunStrategy-Waverun-chevron-down']")
                )
        );
        ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", expandButton1);
        try {
            Thread.sleep(3000);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
        expandButton1.click();
        System.out.println("Wave Run chevron-down button clicked using native click.");

        try {
            Thread.sleep(3000);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }

        WebElement orderPlanningRunInputField = wait.until(
                ExpectedConditions.elementToBeClickable(By.xpath("//ion-input[@data-component-id='OrderPlanningRunId-lookup-dialog-filter-input']//input"))
        );
        if (waveNumber != null && !waveNumber.isEmpty()) {
            orderPlanningRunInputField.click();
            orderPlanningRunInputField.clear();
            orderPlanningRunInputField.sendKeys(waveNumber);
            orderPlanningRunInputField.sendKeys(Keys.ENTER);
        }
        System.out.println("Order Planning Run ID entered: " + waveNumber);
    }

    public static void navigateTillWaveRuns1() throws InterruptedException, IOException {
        WebElement waveRunsLink = driver.findElement(By.cssSelector("a[title='Wave Runs']"));
        ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", waveRunsLink);
        waveRunsLink.click();
        Thread.sleep(5000);
        System.out.println("Wave runs clicked.");
    }

    public static void GenerateTask() throws InterruptedException, IOException {
        System.out.println("GenerateTask START");
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(time));

        Thread.sleep(12000);
        System.out.println("Order Planning Run ID entered: " + waveNumber);

        wait.until(ExpectedConditions.invisibilityOfElementLocated(By.cssSelector("ion-backdrop")));
        JavascriptExecutor js = (JavascriptExecutor) driver;
        System.out.println("Done1");

        try {
            WebElement cardView = wait.until(ExpectedConditions.elementToBeClickable(
                    By.cssSelector("card-view[data-component-id='Card-View'] .card-row.primary[tabindex='0']")
            ));
            js.executeScript("arguments[0].scrollIntoView(true);", cardView);
            cardView.click();
            System.out.println("Card view selected.");
        } catch (StaleElementReferenceException staleEx) {
            System.out.println("Stale element detected. Retrying...");
            WebElement cardViewRetry = wait.until(ExpectedConditions.elementToBeClickable(
                    By.cssSelector("card-view[data-component-id='Card-View'] .card-row.primary[tabindex='0']")
            ));
            js.executeScript("arguments[0].scrollIntoView(true);", cardViewRetry);
            cardViewRetry.click();
            System.out.println("Card view selected after retry.");
        } catch (Exception e) {
            System.out.println("Failed to select the card view: " + e.getMessage());
        }

        System.out.println("Searching Presence of footer-panel-more-label");

        // wait = new WebDriverWait(driver, Duration.ofSeconds(15));
        wait.until(ExpectedConditions.presenceOfElementLocated(By.cssSelector(".footer-panel-more-label")));
        WebElement moreButton = wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("[data-component-id='footer-panel-more-actions']")));
        moreButton.click();


        WebElement generateTaskButton = wait.until(ExpectedConditions.elementToBeClickable(
                By.cssSelector("button[data-component-id='footer-panel-more-actions-GenerateTask']")
        ));
        generateTaskButton.click();


        WebElement submitButton = wait.until(ExpectedConditions.elementToBeClickable(
                By.cssSelector("ion-button[data-component-id='submit-btn']")
        ));
        submitButton.click();

        System.out.println("Generate end");
    }

    public static void clickTasks(XWPFDocument document, String docPathLocal) throws IOException, InterruptedException {
        try {
            Thread.sleep(3000);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
        WebDriverWait waitShort = new WebDriverWait(driver, Duration.ofSeconds(time));
        By allocationsLocator = By.xpath("//ion-item[@data-component-id='Tasks']//a[text()='Tasks']");
        WebElement allocationsLink = waitShort.until(ExpectedConditions.elementToBeClickable(allocationsLocator));
        JavascriptExecutor js = (JavascriptExecutor) driver;
        js.executeScript("arguments[0].click();", allocationsLink);
        System.out.println("Tasks clicked.");

        try {
            Thread.sleep(3000);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }

        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(time));
        WebElement cardPanel = wait.until(ExpectedConditions.visibilityOfElementLocated(
                By.xpath("/html/body/app-root/ion-app/div/ion-split-pane/ion-router-outlet/screen-page/div/div/div[2]/div/ion-content/card-panel")
        ));
        try {
            Thread.sleep(10000);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }

        DocPathManager.captureScreenshot("Tasks",driver);
        DocPathManager.captureAllCardsScreenshots(driver);
         DocPathManager.saveSharedDocument();
    }

    public static void clickorders(XWPFDocument document, String docPathLocal) throws InterruptedException, IOException {
        try {
            Thread.sleep(3000);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
        WebDriverWait waitShort = new WebDriverWait(driver, Duration.ofSeconds(time));
        By allocationsLocator = By.xpath("//ion-item[@data-component-id='Orders']//a[text()='Orders']");
        WebElement allocationsLink = waitShort.until(ExpectedConditions.elementToBeClickable(allocationsLocator));
        JavascriptExecutor js = (JavascriptExecutor) driver;
        js.executeScript("arguments[0].click();", allocationsLink);
        System.out.println("Orders clicked.");

        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(time));
        WebElement cardPanel = wait.until(ExpectedConditions.visibilityOfElementLocated(
                By.xpath("/html/body/app-root/ion-app/div/ion-split-pane/ion-router-outlet/screen-page/div/div/div[2]/div/ion-content/card-panel")
        ));
        try {
            Thread.sleep(10000);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }

        DocPathManager.captureScreenshot("Orders",driver);
        DocPathManager.captureAllCardsScreenshots(driver);
         DocPathManager.saveSharedDocument();
    }

    public static void clickOLPNs(XWPFDocument document, String docPathLocal) throws IOException, InterruptedException {
        try {
            Thread.sleep(3000);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
        WebDriverWait waitShort = new WebDriverWait(driver, Duration.ofSeconds(time));
        By oLPNsLocator = By.xpath("//ion-item[@data-component-id='oLPNs']//a[text()='oLPNs']");
        WebElement oLPNsLink = waitShort.until(ExpectedConditions.elementToBeClickable(oLPNsLocator));
        JavascriptExecutor js = (JavascriptExecutor) driver;
        js.executeScript("arguments[0].click();", oLPNsLink);
        System.out.println("oLPNs clicked.");

        try {
            Thread.sleep(3000);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }

        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(time));
        WebElement cardPanel = wait.until(ExpectedConditions.visibilityOfElementLocated(
                By.xpath("/html/body/app-root/ion-app/div/ion-split-pane/ion-router-outlet/screen-page/div/div/div[2]/div/ion-content/card-panel")
        ));
        try {
            Thread.sleep(10000);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }

        DocPathManager.captureScreenshot("Olpns",driver);
        DocPathManager.captureAllCardsScreenshots(driver);
         DocPathManager.saveSharedDocument();
    }

    // (Old DocumentName kept for compatibility but unused)
    public static void DocumentName(String docName, String filePath) {
        // Deprecated in favor of buildDocPath
        Random rand = new Random();
        int randomNum = rand.nextInt(100000);
        String uniqueDocName = docName + "_" + randomNum;
        String docPath = filePath + uniqueDocName + ".docx";
        System.out.println(docPath);
    }

    public static void ReleaseTask(XWPFDocument document, String docPathLocal) throws InterruptedException, IOException {
        Thread.sleep(5000);
        WebElement selectAllButton = driver.findElement(By.cssSelector("button[data-component-id='selectAllRows']"));
        selectAllButton.click();
        try {
            Thread.sleep(2000);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }

        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(time));
        Thread.sleep(2000);

        WebElement releaseButton = wait.until(
                ExpectedConditions.elementToBeClickable(
                        By.cssSelector("ion-button[data-component-id='footer-panel-action-Release']")
                )
        );
        releaseButton.click();
        Thread.sleep(3000);
        DocPathManager.captureScreenshot("Clicked Release",driver);
         DocPathManager.saveSharedDocument();
        Thread.sleep(10000);

        WebElement yesButton = wait.until(
                ExpectedConditions.elementToBeClickable(
                        By.cssSelector("button[data-component-id='Yes']")
                )
        );
        yesButton.click();
        Thread.sleep(10000);
        System.out.println("Check whether task is released or not");
        DocPathManager.captureScreenshot("Tasks",driver);
        DocPathManager.captureAllCardsScreenshots(driver);
         DocPathManager.saveSharedDocument();

    }

    public static void updateOpsStatus(String filePath, String testcase, String waveStatus) {
        closeExcelIfOpen();
        String result = "Failed";
        if (waveStatus != null && !waveStatus.isEmpty()) {
            result = waveStatus;
        }
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheet("OPS_Tab");
            if (sheet == null) return;
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;
                Cell tcCell = row.getCell(0);
                if (tcCell == null) continue;
                String tc = tcCell.getStringCellValue().trim();
                if (tc.equalsIgnoreCase(testcase.trim())) {
                    Cell resultCell = row.getCell(2);
                    if (resultCell == null) resultCell = row.createCell(2);
                    resultCell.setCellValue(result);
                    break;
                }
            }
            try (FileOutputStream fos = new FileOutputStream(filePath)) {
                workbook.write(fos);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
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



}