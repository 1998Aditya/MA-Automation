//package MA_MSG_Suite_OB;
//
//
//
//import com.google.gson.JsonObject;
//import okhttp3.*;
//import org.apache.poi.ss.usermodel.*;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//import java.io.FileInputStream;
//import java.io.IOException;
//import java.util.*;
//import com.google.gson.JsonParser;
//import com.google.gson.JsonElement;
//import okhttp3.*;
//import org.apache.poi.ss.usermodel.Cell;
//import org.apache.poi.ss.usermodel.Row;
//import org.apache.poi.ss.usermodel.Sheet;
//import org.apache.poi.ss.usermodel.Workbook;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//import org.openqa.selenium.*;
//import java.util.ArrayList;
//import java.util.List;
//
//
//
//
//public class MainA3_oLPNPalletised {
//
//        private static final OkHttpClient client = new OkHttpClient();
//        public static class PalletisedSheetData {
//            String WCSOrderId;
//            String PalletId;
//            String LocationId;
//
//            @Override
//            public String toString() {
//                return "WCSOrderId: " + WCSOrderId + ", PalletId: " + PalletId + ", LocationId: " + LocationId;
//            }
//        }
//
//        public static void main(String filePath, String messageType) {
//            try {
//                List<PalletisedSheetData> palletList = readPalletisedData(filePath);
//                System.out.println("✅ Extracted oLPNPalletised Data:");
//                palletList.forEach(System.out::println);
//
//                String token = getAuthTokenFromExcel();
//                if (token == null) {
//                    System.err.println("❌ Authentication failed.");
//                    return;
//                }
//
//                triggerAPI(palletList, token,filePath,messageType);
//
//            } catch (Exception e) {
//                System.err.println("❌ Error: " + e.getMessage());
//                e.printStackTrace();
//            }
//        }
//
//        private static List<PalletisedSheetData> readPalletisedData(String path) throws IOException {
//            List<PalletisedSheetData> list = new ArrayList<>();
//
//            try (FileInputStream fis = new FileInputStream(path);
//                 Workbook workbook = new XSSFWorkbook(fis)) {
//
//                Sheet sheet = workbook.getSheet("Tasks");
//                if (sheet == null) {
//                    System.err.println("❌ Sheet 'Tasks' not found.");
//                    return list;
//                }
//
//                for (int i = 1; i <= sheet.getLastRowNum(); i++) {
//                    Row row = sheet.getRow(i);
//                    if (row == null) continue;
//
//                    PalletisedSheetData data = new PalletisedSheetData();
//                    data.WCSOrderId = getCellValueAsString(row.getCell(9));   // OLPN
//                    data.PalletId = getCellValueAsString(row.getCell(20));    // PalletId
//                    data.LocationId = getCellValueAsString(row.getCell(21));  // LocationId
//
//                    if (!data.WCSOrderId.isEmpty() && !data.PalletId.isEmpty() && !data.LocationId.isEmpty()) {
//                        list.add(data);
//                    }
//                }
//            }
//            return list;
//        }
//
//        private static JsonObject buildPayload(PalletisedSheetData data) {
//            JsonObject json = new JsonObject();
//            json.addProperty("WCSOrderId", data.WCSOrderId);
//            json.addProperty("PalletId", data.PalletId);
//            json.addProperty("LocationId", data.LocationId);
//            json.addProperty("MessageType", "oLPNPalletised");
//            json.addProperty("UniqueKey", String.valueOf(System.currentTimeMillis()));
//            return json;
//        }
//
//    public static String getAuthTokenFromExcel() throws IOException {
//        ExcelReader reader = new ExcelReader();
//        String LOGIN_URL   = reader.getCellValueByHeader(1, "LOGIN_URL");
//        String UIUsername  = reader.getCellValueByHeader(1, "username");
//        String UIPassword  = reader.getCellValueByHeader(1, "password");
//        reader.close();
//
//        OkHttpClient client = new OkHttpClient();
//        MediaType mediaType = MediaType.parse("application/x-www-form-urlencoded");
//        RequestBody body = RequestBody.create(mediaType,
//                "grant_type=password&username=" + UIUsername + "&password=" + UIPassword);
//
//        Request request = new Request.Builder()
//                .url(LOGIN_URL)
//                .method("POST", body)
//                .addHeader("Content-Type", "application/x-www-form-urlencoded")
//                .addHeader("Authorization", "Basic dWpkc3N0YWdlMTpFYXJ0aC1Nb29uLVN1bjE=")
//                .build();
//
//        Response response = client.newCall(request).execute();
//        String responseBody = response.body() != null ? response.body().string() : null;
//
//        JsonObject json = JsonParser.parseString(responseBody).getAsJsonObject();
//        return json.has("access_token") ? json.get("access_token").getAsString() : null;
//    }
//
//    private static void triggerAPI(List<PalletisedSheetData> palletList, String token,String filePath,String messageType) throws IOException {
//            for (PalletisedSheetData data : palletList) {
//                JsonObject payload = buildPayload(data);
//
//                System.out.println("\n📤 Sending Payload for WCSOrderId: " + data.WCSOrderId);
//                System.out.println(payload.toString());
//
//
//                ExcelReader reader = new ExcelReader();
//                String BASE_URL             = reader.getCellValueByHeader(1, "BASE_URL");
//                String SelectedOrganization = reader.getCellValueByHeader(1, "SelectedOrganization");
//                String SelectedLocation     = reader.getCellValueByHeader(1, "SelectedLocation");
//                reader.close();
//
//
//
//                RequestBody body = RequestBody.create(payload.toString(), MediaType.get("application/json"));
//                Request request = new Request.Builder()
//                        .url(BASE_URL+"/device-integration/api/deviceintegration/process/oLPNPalletised_FER_Src_EP")
//                        .post(body)
//                        .addHeader("Authorization", "Bearer " + token)
//                        .addHeader("Content-Type", "application/json")
//                        .addHeader("SelectedOrganization", SelectedOrganization)
//                        .addHeader("SelectedLocation", SelectedLocation)
//                        .build();
//
//                try (Response response = client.newCall(request).execute()) {
//                    String responseBody = response.body().string();
//                    System.out.println("🔍 Response Code: " + response.code());
//                    System.out.println("🔍 Response Body: " + responseBody);
//
//                    if (response.isSuccessful()) {
//                        System.out.println("✅ Successfully posted oLPNPalletised for WCSOrderId: " + data.WCSOrderId);
//
//                        // 🔍 Validate PalletId
//                        boolean isValid = validatePalletId(data.WCSOrderId, data.PalletId, token,filePath,messageType);
//                        if (!isValid) {
//                            System.err.println("🛑 Halting process due to PalletId mismatch.");
//                            return;
//                        }
//
//                    } else {
//                        System.err.println("❌ Failed for WCSOrderId " + data.WCSOrderId);
//                        return;
//                    }
//                }
//            }
//        }
//
//        private static boolean validatePalletId(String olpnId, String expectedPalletId, String token,String filePath,String messageType) throws IOException {
//            // ✅ Build correct query with quoted OLPN and PalletId condition
//            JsonObject query = new JsonObject();
//            query.addProperty("Query", "OlpnId = '" + olpnId + "' AND PalletId != ''");
//
//            ExcelReader reader = new ExcelReader();
//            String BASE_URL             = reader.getCellValueByHeader(1, "BASE_URL");
//            String SelectedOrganization = reader.getCellValueByHeader(1, "SelectedOrganization");
//            String SelectedLocation     = reader.getCellValueByHeader(1, "SelectedLocation");
//            reader.close();
//
//
//
//            RequestBody body = RequestBody.create(query.toString(), MediaType.get("application/json"));
//            Request request = new Request.Builder()
//                    .url(BASE_URL)
//                    .post(body)
//                    .addHeader("Authorization", "Bearer " + token)
//                    .addHeader("Content-Type", "application/json")
//                    .addHeader("SelectedOrganization", SelectedOrganization)
//                    .addHeader("SelectedLocation", SelectedLocation)
//                    .build();
//
//            try (Response response = client.newCall(request).execute()) {
//                String responseBody = response.body().string();
//                System.out.println("🔍 Validation Response: " + responseBody);
//
//                if (!response.isSuccessful()) {
//                    System.err.println("❌ Validation API failed for OLPN: " + olpnId);
//                    return false;
//                }
//
//                JsonObject json = JsonParser.parseString(responseBody).getAsJsonObject();
//                JsonElement dataElement = json.has("data") ? json.get("data") : null;
//
//                if (dataElement == null || !dataElement.isJsonArray() || dataElement.getAsJsonArray().size() == 0) {
//                    System.err.println("❌ No 'data' found for OLPN: " + olpnId);
//                    return false;
//                }
//
//                JsonObject result = dataElement.getAsJsonArray().get(0).getAsJsonObject();
//                String actualPalletId = result.has("PalletId") ? result.get("PalletId").getAsString() : "";
//
//                boolean match = expectedPalletId.equals(actualPalletId);
//                if (!match) {
//                    System.err.println("❌ PalletId mismatch for OLPN " + olpnId + ". Expected: " + expectedPalletId + ", Found: " + actualPalletId);
//                } else {
//                    System.out.println("✅ OLPN " + olpnId + " has correct PalletId.");
//                }
//                System.out.println("\n⏳ Waiting 20 seconds before launching MHEJournalValidator...");
//                Thread.sleep(20000);
//                System.out.println("🚀 Launching MHEJournalValidator...\n");
//                MainA0_MHEValidator.main(filePath,messageType);
//                return match;
//            } catch (InterruptedException e) {
//                throw new RuntimeException(e);
//            }
//        }
//
//
//        private static String getCellValueAsString(Cell cell) {
//            if (cell == null) return "";
//            switch (cell.getCellType()) {
//                case STRING: return cell.getStringCellValue().trim();
//                case NUMERIC: return String.valueOf((long) cell.getNumericCellValue()).trim();
//                case BOOLEAN: return String.valueOf(cell.getBooleanCellValue()).trim();
//                case FORMULA: return cell.getCellFormula().trim();
//                default: return "";
//            }
//        }
//    }







package MA_MSG_Suite_OB;

import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;
import com.google.gson.stream.JsonReader;

import okhttp3.MediaType;
import okhttp3.OkHttpClient;
import okhttp3.Request;
import okhttp3.Response;
import okhttp3.RequestBody;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.channels.FileChannel;
import java.nio.channels.FileLock;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;
import java.time.Duration;
import java.util.*;
import java.util.stream.Collectors;

/**
 * API-only palletisation using ExcelReader (same style as dummy.txt):
 * 1) Fetch OLPNs from Excel (sheet 'Tasks', columns: Testcase, OLPNs)
 * 2) Batch PickPack /pickpack/api/pickpack/olpn/search with IN (...) to get ShipViaId
 * 3) Group OLPNs by ShipViaId
 * 4) For each group: generate ONE random 10-digit pallet; post one payload per OLPN
 * 5) Write pallet IDs back to Excel ('Tasks', column 'pallet') per OLPN
 */
public class MainA3_oLPNPalletised {
    public static String docPathLocal ;

    // Reuse a single OkHttp client
    private static final OkHttpClient httpClient = new OkHttpClient.Builder()
            .followRedirects(true)
            .followSslRedirects(true)
            .callTimeout(Duration.ofSeconds(30))
            .connectTimeout(Duration.ofSeconds(20))
            .readTimeout(Duration.ofSeconds(30))
            .writeTimeout(Duration.ofSeconds(30))
            .build();

    // --------------------------------------------------------------------
    // Entry points
    // --------------------------------------------------------------------
    public static void main(String filePath, String testcase, String messageType,String env) {
        docPathLocal = DocPathManager.getOrCreateDocPath(filePath, testcase);


        try {
            // 1) Fetch OLPNs
           // testcase="TST_001";
            List<String> olpns = getOlpnsForTestcase(filePath, testcase);
            if (olpns == null || olpns.isEmpty()) {
                System.err.println("❌ No OLPNs found for testcase: " + testcase);
                return;
            }
            System.out.println("📦 OLPNs for '" + testcase + "': " + olpns);

            // 2) Auth
            String token = getAuthTokenFromExcel();
            if (token == null || token.isBlank()) {
                System.err.println("❌ Authentication failed.");
                return;
            }

            // 3) Group by ShipViaId via a single batch call
            Map<String, List<String>> groups = groupOlpnsByShipViaIdBatch(olpns, token);
            printGroups(groups);

            String locationId = "LB02X";

            // 4) API-only post payloads + write back pallets
            palletiseByShipVia(groups, locationId, messageType, token, filePath, testcase);

            System.out.println("✅ Completed palletisation API posting for all ShipVia groups.");

            Thread.sleep(5000);
            Main100_MHEJournalScreenshot.main(testcase,filePath, env, messageType,docPathLocal);
            Thread.sleep(5000);


        } catch (Exception e) {
            System.err.println("❌ Error: " + e.getMessage());
            e.printStackTrace();
        }
    }

//    public static void main(String filePath, String messageType) {
//        main(filePath, testcase, messageType,env);
//    }

    private static void printGroups(Map<String, List<String>> grouped) {
        System.out.println("==================================================");
        System.out.println("📚 Segregated OLPNs by ShipViaId:");
        grouped.forEach((shipVia, list) -> System.out.println("  " + shipVia + " -> " + list));
        System.out.println("==================================================");
    }

    // --------------------------------------------------------------------
    // Excel: Fetch OLPNs for a given Testcase from sheet 'Tasks'
    // --------------------------------------------------------------------
    public static List<String> getOlpnsForTestcase(String filePath, String testcase) {
        List<String> olpns = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            // If needed: testcase = "TST_001";

            Sheet sheet = workbook.getSheet("Tasks");
            if (sheet == null) {
                System.err.println("❌ Sheet 'Tasks' not found in: " + filePath);
                return olpns;
            }
            Row headerRow = sheet.getRow(0);
            if (headerRow == null) {
                System.err.println("❌ Header row missing in 'Tasks'.");
                return olpns;
            }

            int colTestcase = -1, colOlpns = -1;
            for (int c = headerRow.getFirstCellNum(); c < headerRow.getLastCellNum(); c++) {
                String header = getCellString(headerRow.getCell(c, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL));
                if (header == null) continue;
                if (header.trim().equalsIgnoreCase("Testcase")) colTestcase = c;
                else if (header.trim().equalsIgnoreCase("OLPNs")) colOlpns = c;
            }

            if (colTestcase == -1 || colOlpns == -1) {
                System.err.println("❌ Required columns not found. Testcase idx=" + colTestcase + ", OLPNs idx=" + colOlpns);
                return olpns;
            }

            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                String tcValue = getCellString(row.getCell(colTestcase, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL));
                if (tcValue == null) continue;
                if (tcValue.trim().equalsIgnoreCase(testcase.trim())) {
                    String olpnsCell = getCellString(row.getCell(colOlpns, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL));
                    if (olpnsCell != null && !olpnsCell.trim().isEmpty()) {
                        String[] parts = olpnsCell.split("[,;\\n\\r\\t ]+");
                        for (String p : parts) {
                            String t = p.trim();
                            if (!t.isEmpty()) olpns.add(t);
                        }
                    }
                }
            }
        } catch (IOException e) {
            System.err.println("❌ IO Error reading Excel: " + e.getMessage());
        } catch (Exception e) {
            System.err.println("❌ Unexpected error: " + e.getMessage());
        }
        return olpns;
    }

    private static String getCellString(Cell cell) {
        if (cell == null) return null;
        switch (cell.getCellType()) {
            case STRING:
                String s = cell.getStringCellValue();
                return (s == null) ? null : s.trim();
            case NUMERIC:
                DataFormatter df = new DataFormatter();
                return df.formatCellValue(cell).trim();
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue()).trim();
            case FORMULA:
                DataFormatter formatter = new DataFormatter();
                FormulaEvaluator eval = cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
                return formatter.formatCellValue(cell, eval).trim();
            case BLANK:
            case _NONE:
            case ERROR:
            default:
                return null;
        }
    }

    // --------------------------------------------------------------------
    // Auth token using ExcelReader (same as dummy.txt)
    // --------------------------------------------------------------------
    public static String getAuthTokenFromExcel() throws IOException {
        ExcelReader reader = null;
        try {
            reader = new ExcelReader();
            String LOGIN_URL  = reader.getCellValueByHeader(1, "LOGIN_URL");
            String UIUsername = reader.getCellValueByHeader(1, "username");
            String UIPassword = reader.getCellValueByHeader(1, "password");

            requireHttpUrl(LOGIN_URL, "LOGIN_URL");

            MediaType mediaType = MediaType.parse("application/x-www-form-urlencoded");
            RequestBody body = RequestBody.create(mediaType,
                    "grant_type=password&username=" + (UIUsername == null ? "" : UIUsername.trim())
                            + "&password=" + (UIPassword == null ? "" : UIPassword.trim()));

            Request request = new Request.Builder()
                    .url(LOGIN_URL.trim())
                    .method("POST", body)
                    .addHeader("Content-Type", "application/x-www-form-urlencoded")
                    .addHeader("Authorization", "Basic dWpkc3N0YWdlMTpFYXJ0aC1Nb29uLVN1bjE=")
                    .build();

            try (Response response = httpClient.newCall(request).execute()) {
                String responseBody = response.body() != null ? response.body().string() : "";
                if (!response.isSuccessful()) {
                    System.err.println("❌ Auth failed: " + response.code() + " | " + responseBody);
                    return null;
                }
                JsonObject json = JsonParser.parseString(responseBody).getAsJsonObject();
                return json.has("access_token") ? json.get("access_token").getAsString() : null;
            }
        } finally {
            if (reader != null) try { reader.close(); } catch (Exception ignored) {}
        }
    }

    // --------------------------------------------------------------------
    // Batch ShipVia grouping (YOUR implementation, used as-is)
    // --------------------------------------------------------------------
    public static Map<String, List<String>> groupOlpnsByShipViaIdBatch(List<String> olpns, String bearerToken) throws IOException {
        Map<String, List<String>> grouped = new LinkedHashMap<>();
        if (olpns == null || olpns.isEmpty()) {
            System.err.println("⚠️ No OLPNs provided for grouping.");
            return grouped;
        }

        String inList = olpns.stream().filter(s -> s != null && !s.isBlank()).map(String::trim).collect(Collectors.joining(","));
        String body = "{\"Query\":\"OlpnId in (" + inList + ")\"}";
        System.out.println("📤 Batch body: " + body);

        ExcelReader reader = new ExcelReader();
        String BASE_URL = reader.getCellValueByHeader(1, "BASE_URL");
        String SelectedOrganization = reader.getCellValueByHeader(1, "SelectedOrganization");
        String SelectedLocation = reader.getCellValueByHeader(1, "SelectedLocation");
        reader.close();

        requireHttpUrl(BASE_URL, "BASE_URL");
        String url = joinUrl(BASE_URL, "/pickpack/api/pickpack/olpn/search");

        RequestBody requestBody = RequestBody.create(MediaType.parse("application/json"), body);
        Request.Builder rb = new Request.Builder()
                .url(url)
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
        Map<String, String> olpnShipViaMap = new HashMap<>();
        try (Response response = httpClient.newCall(request).execute()) {
            String responseBody = response.body() != null ? response.body().string() : "";
            String contentType = Optional.ofNullable(response.header("Content-Type")).orElse("");

            if (response.code() >= 200 && response.code() < 300 && contentType.toLowerCase(Locale.ROOT).contains("application/json")) {
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
                            String shipVia = readString(item, "ShipViaId");
                            if (isBlank(shipVia)) shipVia = readString(item, "FinalDeliveryShipViaId");
                            if (isBlank(shipVia)) shipVia = readString(item, "ServiceLevelId");
                            if (isBlank(shipVia)) shipVia = readString(item, "StaticRouteId");
                            if (isBlank(shipVia)) {
                                for (Map.Entry<String, JsonElement> e : item.entrySet()) {
                                    String key = e.getKey();
                                    if (key.toLowerCase(Locale.ROOT).contains("shipviaid") && !e.getValue().isJsonNull()) {
                                        String val = e.getValue().getAsString();
                                        if (!isBlank(val)) {
                                            shipVia = val.trim();
                                            break;
                                        }
                                    }
                                }
                            }
                            if (isBlank(shipVia)) shipVia = "UNKNOWN";
                            olpnShipViaMap.put(olpn, shipVia);
                        }
                    }
                }
            }
        } catch (Exception e) {
            System.err.println("❌ Batch request failed: " + e);
        }

        for (String id : olpns) {
            String shipVia = olpnShipViaMap.getOrDefault(id, "UNKNOWN");
            grouped.computeIfAbsent(shipVia, k -> new ArrayList<>()).add(id);
            System.out.println("↪️ OLPN " + id + " → ShipViaId: " + shipVia);
        }
        return grouped;
    }

    // --------------------------------------------------------------------
    // Orchestrate: one pallet per ShipVia group (API-only) + write back
    // --------------------------------------------------------------------
    private static void palletiseByShipVia(Map<String, List<String>> groups,
                                           String locationId,
                                           String messageType,
                                           String token,
                                           String filePath,
                                           String testcase) throws IOException {

        Map<String, String> olpnToPalletUsed = new LinkedHashMap<>();

        for (Map.Entry<String, List<String>> e : groups.entrySet()) {
            String shipVia = e.getKey();
            List<String> olpnsInGroup = e.getValue();
            if (olpnsInGroup == null || olpnsInGroup.isEmpty()) continue;

            // 1) Generate one pallet for this ShipVia group
            String palletId = generateNumericPalletId10();
            System.out.println("📦 Using PalletId " + palletId + " for ShipVia: " + shipVia);

            // 2) Post payload per OLPN with same pallet & record used pallet
            for (String olpn : olpnsInGroup) {
                if (olpn == null || olpn.isBlank()) continue;

                JsonObject payload = new JsonObject();
                payload.addProperty("WCSOrderId", olpn.trim());
                payload.addProperty("PalletId", palletId);
                payload.addProperty("LocationId", locationId);
                payload.addProperty("MessageType", messageType);
                payload.addProperty("UniqueKey", uniqueKey());

                boolean posted = postDeviceIntegration(payload, token);
                if (!posted) {
                    System.err.println("❌ Posting failed for OLPN: " + olpn + " (ShipVia: " + shipVia + ")");
                } else {
                    System.out.println("✅ Posted payload for OLPN " + olpn + " (ShipVia: " + shipVia + ")");
                }

                olpnToPalletUsed.put(olpn.trim(), palletId);
            }
        }
        closeExcelIfOpen();
        // 3) Write back pallet IDs into Tasks sheet (column "pallet") per OLPN
        writePalletsToExcel(filePath, testcase, olpnToPalletUsed);


    }

    // --------------------------------------------------------------------
    // Device-integration POST per payload (API-only, ExcelReader style)
    // --------------------------------------------------------------------
    private static boolean postDeviceIntegration(JsonObject payload, String token) throws IOException {
        ExcelReader reader = new ExcelReader();
        String BASE_URL             = reader.getCellValueByHeader(1, "BASE_URL");
        String SelectedOrganization = reader.getCellValueByHeader(1, "SelectedOrganization");
        String SelectedLocation     = reader.getCellValueByHeader(1, "SelectedLocation");
        reader.close();

        requireHttpUrl(BASE_URL, "BASE_URL");
        String url = joinUrl(BASE_URL, "/device-integration/api/deviceintegration/process/oLPNPalletised_FER_Src_EP");

        RequestBody body = RequestBody.create(payload.toString(), MediaType.get("application/json"));
        Request request = new Request.Builder()
                .url(url)
                .post(body)
                .addHeader("Authorization", bearer(token))
                .addHeader("Content-Type", "application/json")
                .addHeader("SelectedOrganization", SelectedOrganization)
                .addHeader("SelectedLocation", SelectedLocation)
                .build();

        try (Response response = httpClient.newCall(request).execute()) {
            String responseBody = response.body() != null ? response.body().string() : "";
            System.out.println("🔍 Post Response Code: " + response.code());
            System.out.println("🔍 Post Response Body: " + responseBody);
            if (!response.isSuccessful()) {
                System.err.println("❌ Device-integration post failed: HTTP " + response.code());
                return false;
            }
            return true;
        }
    }

    // --------------------------------------------------------------------
    // NEW: Wait for exclusive access & write pallet IDs back to Excel
    // --------------------------------------------------------------------
    private static void writePalletsToExcel(String filePath,
                                            String testcase,
                                            Map<String, String> olpnToPallet) throws IOException {
        if (olpnToPallet == null || olpnToPallet.isEmpty()) {
            System.out.println("ℹ️ No pallets to write back to Excel.");
            return;
        }

        // 1) Wait until the file is writable (not locked by Excel/another process)
        boolean ready = waitForExclusiveAccess(filePath, 30_000, 500);
        if (!ready) {
            throw new IOException(filePath + " is locked. Close the Excel file and retry.");
        }

        // 2) Read workbook (input phase) — close stream before writing
        XSSFWorkbook workbook;
        try (FileInputStream fis = new FileInputStream(filePath)) {
            workbook = new XSSFWorkbook(fis);
        } // fis closed here

        try {
            Sheet sheet = workbook.getSheet("Tasks");
            if (sheet == null) {
                System.err.println("⚠️ Sheet 'Tasks' not found while writing pallets.");
                workbook.close();
                return;
            }

            Row headerRow = sheet.getRow(0);
            if (headerRow == null) {
                System.err.println("⚠️ Header row missing while writing pallets.");
                workbook.close();
                return;
            }

            int colTestcase = -1, colOlpns = -1, colPallet = -1;

            for (int c = headerRow.getFirstCellNum(); c < headerRow.getLastCellNum(); c++) {
                Cell cell = headerRow.getCell(c, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                String header = getCellString(cell);
                if (header == null) continue;
                String h = header.trim();
                if (h.equalsIgnoreCase("Testcase")) colTestcase = c;
                else if (h.equalsIgnoreCase("OLPNs")) colOlpns = c;
                else if (h.equalsIgnoreCase("pallet")) colPallet = c;
            }

            // Create 'pallet' column if missing
            if (colPallet == -1) {
                colPallet = headerRow.getLastCellNum();
                if (colPallet < 0) colPallet = 0;
                Cell newHeader = headerRow.createCell(colPallet, CellType.STRING);
                newHeader.setCellValue("pallet");
                System.out.println("🆕 Created 'pallet' column at index " + colPallet);
            }

            if (colTestcase == -1 || colOlpns == -1) {
                System.err.println("⚠️ Required columns not found while writing pallets. Testcase idx=" +
                        colTestcase + ", OLPNs idx=" + colOlpns);
                workbook.close();
                return;
            }

            DataFormatter df = new DataFormatter();
            FormulaEvaluator eval = workbook.getCreationHelper().createFormulaEvaluator();

            int updates = 0;

            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;

                String tcVal = getCellString(row.getCell(colTestcase, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL));
                if (tcVal == null || !tcVal.trim().equalsIgnoreCase(testcase.trim())) continue;

                String olpnVal = df.formatCellValue(row.getCell(colOlpns, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL), eval).trim();
                if (olpnVal.isEmpty()) continue;

                String pallet = olpnToPallet.get(olpnVal);
                if (pallet == null || pallet.trim().isEmpty()) continue;

                Cell palletCell = row.getCell(colPallet, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                palletCell.setCellType(CellType.STRING);
                palletCell.setCellValue(pallet.trim());
                updates++;
            }

            // 3) Write workbook (output phase) — separate stream
            try (FileOutputStream fos = new FileOutputStream(filePath)) {
                workbook.write(fos);
                fos.flush();
            }
            workbook.close();
            System.out.println("✍️ Pallet write-back completed. Rows updated: " + updates);

        } catch (IOException ioe) {
            try { workbook.close(); } catch (Exception ignored) {}
            System.err.println("❌ Error writing pallets to Excel: " + ioe.getMessage());
            throw ioe;
        } catch (Exception e) {
            try { workbook.close(); } catch (Exception ignored) {}
            System.err.println("❌ Unexpected error writing pallets: " + e.getMessage());
        }
    }

    /**
     * Tries to obtain exclusive access to the file by probing a write lock.
     * Retries until timeout. Returns true if writable, false otherwise.
     */
    private static boolean waitForExclusiveAccess(String filePath, long timeoutMillis, long pollMillis) {
        long deadline = System.currentTimeMillis() + timeoutMillis;
        while (System.currentTimeMillis() < deadline) {
            if (canLockForWrite(filePath)) {
                return true;
            }
            try {
                Thread.sleep(Math.max(100, pollMillis));
            } catch (InterruptedException ie) {
                Thread.currentThread().interrupt();
                return false;
            }
        }
        return false;
    }

    /**
     * Attempts to acquire a file lock for write. Releases immediately if successful.
     * Returns true if lock was obtained, false if locked by another process.
     */
    private static boolean canLockForWrite(String filePath) {
        try (FileChannel ch = FileChannel.open(Paths.get(filePath),
                StandardOpenOption.WRITE,
                StandardOpenOption.CREATE)) {
            try (FileLock lock = ch.tryLock()) {
                if (lock != null) {
                    // We got the lock; release immediately
                    return true;
                }
                return false;
            }
        } catch (Exception ex) {
            // If lock cannot be acquired due to sharing violation, we treat as locked
            return false;
        }
    }

    // --------------------------------------------------------------------
    // Helpers
    // --------------------------------------------------------------------
    private static boolean isBlank(String s) { return s == null || s.trim().isEmpty(); }
    private static String readString(JsonObject obj, String key) {
        return (obj.has(key) && !obj.get(key).isJsonNull()) ? obj.get(key).getAsString().trim() : null;
    }

    private static String bearer(String token) {
        String t = token == null ? "" : token.trim();
        return t.toLowerCase(Locale.ROOT).startsWith("bearer ") ? t : "Bearer " + t;
    }

    /** UniqueKey = millis + 8 random digits */
    private static String uniqueKey() {
        long ts = System.currentTimeMillis();
        int rand = new java.security.SecureRandom().nextInt(100_000_000);
        return String.valueOf(ts) + String.format("%08d", rand);
    }

    /** Generate random 10-digit pallet id (first digit non-zero) */
    private static String generateNumericPalletId10() {
        java.security.SecureRandom rnd = new java.security.SecureRandom();
        StringBuilder sb = new StringBuilder(10);
        sb.append(1 + rnd.nextInt(9));
        for (int i = 1; i < 10; i++) sb.append(rnd.nextInt(10));
        return sb.toString();
    }

    private static void requireHttpUrl(String baseUrl, String name) {
        if (baseUrl == null || baseUrl.trim().isEmpty())
            throw new IllegalArgumentException(name + " is missing or blank in Excel (row 1).");
        String trimmed = baseUrl.trim();
        if (!trimmed.startsWith("http://") && !trimmed.startsWith("https://"))
            throw new IllegalArgumentException(name + " must start with http:// or https://. Found: " + baseUrl);
    }

    private static String joinUrl(String base, String path) {
        base = base.trim();
        while (base.endsWith("/")) base = base.substring(0, base.length() - 1);
        if (!path.startsWith("/")) path = "/" + path;
        return base + path;
    }

    // ✅ Close Excel if open
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
