package MA_MSG_Suite_OB;



import com.google.gson.*;
import okhttp3.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;
import com.google.gson.*;
import io.github.bonigarcia.wdm.WebDriverManager;
import okhttp3.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
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
import java.util.Collections;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class MainA0_MHEValidator {

private static final OkHttpClient client = new OkHttpClient();


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




    // 🔧 Change this per class context
     //   private static final String MESSAGE_TYPE = "OrderReadyForPacking,oLPNPrepared,oLPNPalletised,PalletReady"; // Or OLPNPrepared, OLPNPalletised, PalletReady

        public static void main(String EXCEL_PATH,String MESSAGE_TYPE ) {
            try {
                List<String> olpnIds = readOLPNsFromExcel(EXCEL_PATH);
                System.out.println("📄 Extracted OLPNs:");
                olpnIds.forEach(System.out::println);

                String token = getAuthTokenFromExcel();
                if (token == null) {
                    System.err.println("❌ Failed to authenticate.");
                    return;
                }


                for (String olpn : olpnIds) {
                    JsonObject requestBody = buildJournalQuery(olpn, MESSAGE_TYPE);
                    postJournalQuery(olpn, requestBody, token);
                }

            } catch (Exception e) {
                System.err.println("❌ Error: " + e.getMessage());
                e.printStackTrace();
            }
        }

        static List<String> readOLPNsFromExcel(String path) throws IOException {
            List<String> olpnList = new ArrayList<>();

            try (FileInputStream fis = new FileInputStream(path);
                 Workbook workbook = new XSSFWorkbook(fis)) {

                Sheet sheet = workbook.getSheet("Tasks");
                if (sheet == null) {
                    System.err.println("❌ Sheet 'Tasks' not found.");
                    return olpnList;
                }

                for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                    Row row = sheet.getRow(i);
                    if (row == null) continue;

                    String olpn = getCellValueAsString(row.getCell(4));
                    if (!olpn.isEmpty()) {
                        olpnList.add(olpn);
                    }
                }
            }
            return olpnList;
        }

        private static JsonObject buildJournalQuery(String olpnId, String messageType) {
            JsonObject root = new JsonObject();
            root.addProperty("ViewName", "MessageJournal");

            JsonArray filters = new JsonArray();

            JsonObject olpnFilter = new JsonObject();
            olpnFilter.addProperty("ViewName", "MessageJournal");
            olpnFilter.addProperty("AttributeId", "Stage1.MessagePayload");
            olpnFilter.addProperty("Operator", "=");
            olpnFilter.add("FilterValues", new Gson().toJsonTree(Collections.singletonList(olpnId)));
            olpnFilter.addProperty("requiredFilter", false);
            olpnFilter.addProperty("negativeFilter", false);
            filters.add(olpnFilter);

            JsonObject typeFilter = new JsonObject();
            typeFilter.addProperty("ViewName", "MessageJournal");
            typeFilter.addProperty("AttributeId", "MessageType");
            typeFilter.addProperty("Operator", "=");
            typeFilter.add("FilterValues", new Gson().toJsonTree(Collections.singletonList(messageType)));
            typeFilter.addProperty("requiredFilter", false);
            typeFilter.addProperty("negativeFilter", false);
            filters.add(typeFilter);

            root.add("Filters", filters);
            root.add("RequestAttributeIds", new JsonArray());
            root.add("SearchOptions", new JsonArray());
            root.add("SearchChains", new JsonArray());
            root.add("FilterExpression", JsonNull.INSTANCE);
            root.addProperty("Page", 0);
            root.addProperty("TotalCount", -1);
            root.addProperty("SortOrder", "desc");
            root.addProperty("SortIndicator", "chevron-up");
            root.addProperty("TimeZone", "Europe/Paris");
            root.addProperty("IsCommonUI", false);
            root.add("ComponentShortName", JsonNull.INSTANCE);
            root.addProperty("EnableMaxCountLimit", true);
            root.addProperty("MaxCountLimit", 500);
            root.add("PageQuery", JsonNull.INSTANCE);
            root.add("ChildQuery", JsonNull.INSTANCE);
            root.addProperty("ComponentName", "com-manh-cp-dmui-search");
            root.addProperty("Size", 25);
            root.addProperty("AdvancedFilter", false);
            root.addProperty("Sort", "RAW_IN_TIMESTAMP");

            return root;
        }

        private static void postJournalQuery(String olpnId, JsonObject bodyJson, String token) throws IOException {
            ExcelReader reader = new ExcelReader();
            String BASE_URL             = reader.getCellValueByHeader(1, "BASE_URL");
            String SelectedOrganization = reader.getCellValueByHeader(1, "SelectedOrganization");
            String SelectedLocation     = reader.getCellValueByHeader(1, "SelectedLocation");
            reader.close();



            RequestBody body = RequestBody.create(bodyJson.toString(), MediaType.get("application/json"));
            Request request = new Request.Builder()
                    .url(BASE_URL + "/dmui-facade/api/dmui-facade/entity/search")
                    .post(body)
                    .addHeader("Authorization", "Bearer " + token)
                    .addHeader("Content-Type", "application/json")
                    .addHeader("SelectedOrganization", SelectedOrganization)
                    .addHeader("SelectedLocation", SelectedLocation)
                    .build();



            try (Response response = client.newCall(request).execute()) {
                String responseBody = response.body().string();

                if (!response.isSuccessful()) {
                    System.err.println("❌ Journal query failed for OLPN " + olpnId + ": " + response.code());
                    System.err.println("Response: " + responseBody);
                    System.exit(1);
                }

                JsonObject json = JsonParser.parseString(responseBody).getAsJsonObject();
                JsonArray results = json.has("data") && json.getAsJsonObject("data").has("Results")
                        ? json.getAsJsonObject("data").getAsJsonArray("Results")
                        : null;

                if (results == null || results.size() == 0) {
                    System.err.println("⛔ OLPN " + olpnId + " → Status: NOT FOUND");
                    System.exit(1);
                }

                JsonObject result = results.get(0).getAsJsonObject();
                String status = result.has("Status") ? result.get("Status").getAsString().toUpperCase() : "UNKNOWN";
                String messageType = result.has("MessageType") ? result.get("MessageType").getAsString() : "UNKNOWN";

                System.out.println("🔍 OLPN " + olpnId + " → Status: " + status + " | MessageType: " + messageType);

                if (status.equals("FAILED")) {
                    System.err.println("⛔ OLPN " + olpnId + " failed. Stopping process.");
                    System.exit(1);
                } else if (!status.equals("COMPLETED") && !status.equals("IN PROGRESS")) {
                    System.err.println("⚠️ OLPN " + olpnId + " → Unexpected Status: " + status);
                    System.exit(1);
                }
            }
        }


//        private static String getAccessToken() throws IOException {
//            RequestBody body = new FormBody.Builder()
//                    .add("grant_type", "password")
//                    .add("username", USERNAME)
//                    .add("password", PASSWORD)
//                    .add("client_id", CLIENT_ID)
//                    .add("client_secret", CLIENT_SECRET)
//                    .build();
//
//            Request request = new Request.Builder()
//                    .url(TOKEN_URL)
//                    .post(body)
//                    .header("Content-Type", "application/x-www-form-urlencoded")
//                    .build();
//
//            try (Response response = client.newCall(request).execute()) {
//                if (!response.isSuccessful()) {
//                    System.err.println("❌ Authentication failed: " + response.code());
//                    return null;
//                }
//
//                String responseBody = response.body().string();
//                JsonObject json = JsonParser.parseString(responseBody).getAsJsonObject();
//                return json.get("access_token").getAsString();
//            }
//        }

        private static String getCellValueAsString(Cell cell) {
            if (cell == null) return "";
            try {
                switch (cell.getCellType()) {
                    case STRING:
                        return cell.getStringCellValue().trim();
                    case NUMERIC:
                        return String.valueOf((int) cell.getNumericCellValue()).trim();
                    case BOOLEAN:
                        return String.valueOf(cell.getBooleanCellValue()).trim();
                    case FORMULA:
                        switch (cell.getCachedFormulaResultType()) {
                            case STRING:
                                return cell.getStringCellValue().trim();
                            case NUMERIC:
                                return String.valueOf((int) cell.getNumericCellValue()).trim();
                            case BOOLEAN:
                                return String.valueOf(cell.getBooleanCellValue()).trim();
                            default:
                        }
                }
            } finally {

            }
            return "";
        }

}