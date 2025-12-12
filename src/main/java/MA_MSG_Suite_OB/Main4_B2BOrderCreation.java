//
//package MA_MSG_Suite_OB;
//
//import com.google.gson.JsonArray;
//import com.google.gson.JsonElement;
//import com.google.gson.JsonObject;
//import com.google.gson.JsonParser;
//import okhttp3.*;
//import org.apache.poi.ss.usermodel.*;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//
//import java.io.*;
//import java.util.HashMap;
//import java.util.HashSet;
//import java.util.Map;
//import java.util.Set;
//
//public class Main4_B2BOrderCreation {
//
//    public static void main(String testcaseToRun, String filePath) throws IOException {
//        closeExcelIfOpen();
//        FileInputStream fis = null;
//        Workbook workbook = null;
//
//        try {
//            String token = getAuthTokenFromExcel();
//            if (token == null) {
//                System.err.println("Failed to retrieve access token.");
//                return;
//            }
//
//            fis = new FileInputStream(filePath);
//            workbook = new XSSFWorkbook(fis);
//            Sheet sheet = workbook.getSheet("Order Creation");
//
//            JsonObject root = new JsonObject();
//            JsonArray dataArray = new JsonArray();
//
//            // Map OriginalOrderId → Row
//            Map<String, Row> orderIdToRowMap = new HashMap<>();
//
//            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
//                Row row = sheet.getRow(i);
//                if (row == null) continue;
//
//                String testcase = getString(row, 0);
//                if (!testcase.equalsIgnoreCase(testcaseToRun)) continue;
//
//                JsonObject order = new JsonObject();
//                order.addProperty("Cancelled", false);
//                order.addProperty("DeliveryEndDateTime", getString(row, 2));
//                order.addProperty("DeliveryStartDateTime", getString(row, 3));
//                order.addProperty("DestinationFacilityId", getString(row, 4));
//
////                String destinationRaw = getString(row, 3);
////                JsonObject destinationJson = JsonParser.parseString("{" + destinationRaw + "}").getAsJsonObject();
////                order.add("DestinationAddress", destinationJson);
//
//
//                JsonObject extended = new JsonObject();
//                extended.addProperty("MAUJDSDeliveryType", getString(row, 5));
//                extended.addProperty("MAUJDSPrepaid", getString(row, 6));
//                extended.addProperty("MAUJDSPaymentType1Value", getString(row, 7));
//                extended.addProperty("MAUJDSPaymentType2Value", getString(row, 8));
//                extended.addProperty("MAUJDSPaymentType3Value", getString(row, 9));
//                extended.addProperty("MAUJDSPostageCharge", getInt(row, 10));
//                extended.addProperty("MAUJDSSignatory", "");
//                extended.add("MAUJDSPaymentStatus", null);
//                extended.add("MAUJDSPaymentType1", null);
//                extended.add("MAUJDSPaymentType2", null);
//                extended.add("MAUJDSPaymentType3", null);
//                extended.addProperty("MAUJDSFasciaCode", getString(row, 16));
//                extended.addProperty("MAUJDSFasciaDescription", getString(row, 17));
//                extended.addProperty("MAUJDSWebOrderNumber", getString(row, 18));
//                extended.addProperty("MAUJDSTotalPrice", getInt(row, 19));
//                extended.addProperty("MAUJDSPODType", getString(row, 20));
//                extended.addProperty("MAUJDSServicePool", getString(row, 21));
//                extended.addProperty("MAUJDSShipmentType", getString(row, 22));
//                extended.addProperty("MAUJDSDespatchType", getString(row, 23));
//                extended.addProperty("MAUJDSDeliveryInstructions1", "");
//                extended.addProperty("MAUJDSDeliveryInstructions2", "");
//                extended.addProperty("MAUJDSBranchNumber", getString(row, 26));
//                extended.addProperty("MAUJDSSubOrderType", getString(row, 27));
//
//                order.add("Extended", extended);
//                order.addProperty("IncotermId", getString(row, 28));
//                order.addProperty("OrderPlacedDateTime", getString(row, 29));
//                order.addProperty("OrderType", getString(row, 30));
//                order.addProperty("OriginFacilityId", getString(row, 31));
//
//                JsonArray contacts = new JsonArray();
//                JsonObject contactWrapper = new JsonObject();
//                contactWrapper.addProperty("ContactTypeId", getString(row, 32));
//
//                JsonObject contact = new JsonObject();
//                contact.addProperty("FirstName", getString(row, 33));
//                contact.addProperty("LastName", getString(row, 34));
//
//                contactWrapper.add("Contact", contact);
//                contacts.add(contactWrapper);
//                order.add("OriginalOrderContact", contacts);
//
//                String orderId = getString(row, 35);
//                order.addProperty("OriginalOrderId", orderId);
//
//                JsonArray orderLines = new JsonArray();
//                JsonObject line = new JsonObject();
//                line.addProperty("ItemId", getString(row, 36));
//                line.addProperty("OrderedQuantity", getInt(row, 37));
//                line.addProperty("OriginalOrderLineId", getString(row, 38));
//                line.addProperty("QuantityUomId", getString(row, 39));
//                line.addProperty("UnitPrice", getInt(row, 40));
//                orderLines.add(line);
//                order.add("OriginalOrderLine", orderLines);
//
//                order.addProperty("PickupStartDateTime", getString(row, 41));
//                order.addProperty("PreShipConfirmRequired", true);
//
//                dataArray.add(order);
//                orderIdToRowMap.put(orderId.trim(), row);
//            }
//
//            root.add("Data", dataArray);
//            String jsonBody = root.toString();
//
//            String response = callPOST(jsonBody, token);
//            System.out.println("Request Body:\n" + jsonBody);
//            System.out.println("Response:\n" + response);
//
//            // Parse response
//            JsonObject respJson = JsonParser.parseString(response).getAsJsonObject();
//            Set<String> failedOrderIds = new HashSet<>();
//
//            if (respJson.has("data")) {
//                JsonObject data = respJson.getAsJsonObject("data");
//                if (data.has("FailedRecords") && data.get("FailedRecords").isJsonArray()) {
//                    JsonArray failedRecords = data.getAsJsonArray("FailedRecords");
//                    for (JsonElement elem : failedRecords) {
//                        JsonObject failed = elem.getAsJsonObject();
//                        if (failed.has("OriginalOrderId")) {
//                            String failedId = failed.get("OriginalOrderId").getAsString().trim();
//                            failedOrderIds.add(failedId);
//                        }
//                    }
//                }
//            }
//
//            // Mark each row
//            for (Map.Entry<String, Row> entry : orderIdToRowMap.entrySet()) {
//                String orderId = entry.getKey();
//                Row row = entry.getValue();
//                Cell resultCell = row.getCell(1); // Column B
//                if (resultCell == null) resultCell = row.createCell(1, CellType.STRING);
//                resultCell.setCellValue(failedOrderIds.contains(orderId) ? "Failed" : "Passed");
//            }
//
//            fis.close();
//            FileOutputStream fos = new FileOutputStream(filePath);
//            workbook.write(fos);
//            fos.close();
//            workbook.close();
//
//        } catch (Exception e) {
//            System.err.println("❌ Error in execution: " + e.getMessage());
//            e.printStackTrace();
//        }
//    }
//
//
//
//
//    public static String getAuthTokenFromExcel() throws IOException {
//        ExcelReader reader = new ExcelReader();
//
//// By header name
//        String LOGIN_URL = reader.getCellValueByHeader(1, "LOGIN_URL");
//        String UIUsername = reader.getCellValueByHeader(1, "username");
//        String UIPassword = reader.getCellValueByHeader(1, "password");
//
//        reader.close();
//
//
//        // Step 2: Call token API
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
//        JsonObject json = JsonParser.parseString(responseBody).getAsJsonObject();
//
//        return json.has("access_token") ? json.get("access_token").getAsString() : null;
//    }
//
//    public static String callPOST(String body, String token) throws IOException {
//        ExcelReader reader = new ExcelReader();
//
//// By header name
//        String BASE_URL = reader.getCellValueByHeader(1, "BASE_URL");
//        String SelectedOrganization = reader.getCellValueByHeader(1, "SelectedOrganization");
//        String SelectedLocation = reader.getCellValueByHeader(1, "SelectedLocation");
//
//        reader.close();
//
//
//
//        OkHttpClient client = new OkHttpClient();
//        MediaType mediaType = MediaType.parse("application/json");
//        RequestBody requestBody = RequestBody.create(mediaType, body);
//
//        Request request = new Request.Builder()
//                .url(BASE_URL+"/dcorder/api/dcorder/originalOrder/bulkImport")
//                .method("POST", requestBody)
//                .addHeader("Content-Type", "application/json")
//                .addHeader("SelectedOrganization", SelectedOrganization)
//                .addHeader("SelectedLocation", SelectedLocation)
//                .addHeader("Authorization", "Bearer " + token)
//                .build();
//
//        Response response = client.newCall(request).execute();
//        return response.body() != null ? response.body().string() : null;
//    }
//
//    private static String getString(Row row, int index) {
//        Cell cell = row.getCell(index);
//        if (cell == null) return "";
//
//        switch (cell.getCellType()) {
//            case STRING:
//                return cell.getStringCellValue().trim();
//            case NUMERIC:
//                return String.valueOf(cell.getNumericCellValue());
//            case BOOLEAN:
//                return String.valueOf(cell.getBooleanCellValue());
//            case FORMULA:
//                try {
//                    return cell.getStringCellValue().trim();
//                } catch (IllegalStateException e) {
//                    return String.valueOf(cell.getNumericCellValue());
//                }
//            default:
//                return "";
//        }
//    }
//
//    private static int getInt(Row row, int index) {
//        Cell cell = row.getCell(index);
//        if (cell == null) return 0;
//
//        switch (cell.getCellType()) {
//            case NUMERIC:
//                return (int) cell.getNumericCellValue();
//            case STRING:
//                try {
//                    return Integer.parseInt(cell.getStringCellValue().trim());
//                } catch (NumberFormatException e) {
//                    return 0;
//                }
//            case FORMULA:
//                try {
//                    return (int) cell.getNumericCellValue(); // assumes formula returns numeric
//                } catch (Exception e) {
//                    return 0;
//                }
//            default:
//                return 0;
//        }
//    }
//    private static void closeExcelIfOpen() {
//        try {
//            Process process = Runtime.getRuntime().exec("tasklist");
//            BufferedReader reader = new BufferedReader(new InputStreamReader(process.getInputStream()));
//            String line;
//            boolean excelRunning = false;
//            while ((line = reader.readLine()) != null) {
//                if (line.toLowerCase().contains("excel.exe")) {
//                    excelRunning = true;
//                    break;
//                }
//            }


//            if (excelRunning) {
//                System.out.println("⚠️ Excel is open. Closing it...");
//                Runtime.getRuntime().exec("taskkill /IM excel.exe /F");
//                Thread.sleep(2000);
//            }
//        } catch (Exception e) {
//            System.err.println("⚠️ Could not check/close Excel: " + e.getMessage());
//        }
//    }
//}











package MA_MSG_Suite_OB;

import com.google.gson.*;
import okhttp3.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Main4_B2BOrderCreation {

    // ========= Entry point =========
    public static void main(String testcaseToRun, String filePath) throws IOException {
        closeExcelIfOpen();
        FileInputStream fis = null;
        Workbook workbook = null;

        try {
            // --- Token ---
            String token = getAuthTokenFromExcel();
            if (token == null) {
                System.err.println("Failed to retrieve access token.");
                return;
            }

            // --- Open workbook/sheet ---
            fis = new FileInputStream(filePath);
            workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheet("Order Creation");
            if (sheet == null) throw new IllegalStateException("Sheet 'Order Creation' not found.");

            // --- Build header map ---
            Map<String, Integer> headerIndex = buildHeaderIndexMap(sheet);

            // --- Build payload ---
            JsonArray dataArray = new JsonArray();
            Map<String, Row> orderIdToRowMap = new HashMap<>();

            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;

                String testcase = getString(row, headerIndex, "Testcase");
                if (!testcase.equalsIgnoreCase(testcaseToRun)) continue;

                JsonObject order = new JsonObject();

                // Scalars
                boolean cancelled = getBoolean(row, headerIndex, "Cancelled", false);
                order.addProperty("Cancelled", cancelled);
                order.addProperty("DeliveryEndDateTime", getString(row, headerIndex, "DeliveryEndDateTime"));
                order.addProperty("DeliveryStartDateTime", getString(row, headerIndex, "DeliveryStartDateTime"));

                // Destination (Address JSON or Facility Id or Alias)
                addDestination(row, headerIndex, order);
                validateDestination(order);

                // Bill-To (optional)
                addBillTo(row, headerIndex, order);

                // Extended
                JsonObject extended = new JsonObject();
                extended.addProperty("MAUJDSDeliveryType", getString(row, headerIndex, "Extended_MAUJDSDeliveryType"));
                extended.addProperty("MAUJDSPrepaid", getString(row, headerIndex, "Extended_MAUJDSPrepaid"));
                extended.addProperty("MAUJDSPaymentType1Value", getString(row, headerIndex, "Extended_MAUJDSPaymentType1Value"));
                extended.addProperty("MAUJDSPaymentType2Value", getString(row, headerIndex, "Extended_MAUJDSPaymentType2Value"));
                extended.addProperty("MAUJDSPaymentType3Value", getString(row, headerIndex, "Extended_MAUJDSPaymentType3Value"));
                extended.addProperty("MAUJDSPostageCharge", getInt(row, headerIndex, "Extended_MAUJDSPostageCharge"));
                extended.addProperty("MAUJDSFasciaCode", getString(row, headerIndex, "Extended_MAUJDSFasciaCode"));
                extended.addProperty("MAUJDSFasciaDescription", getString(row, headerIndex, "Extended_MAUJDSFasciaDescription"));
                extended.addProperty("MAUJDSWebOrderNumber", getString(row, headerIndex, "Extended_MAUJDSWebOrderNumber"));
                extended.addProperty("MAUJDSTotalPrice", getInt(row, headerIndex, "Extended_MAUJDSTotalPrice"));
                extended.addProperty("MAUJDSServicePool", getString(row, headerIndex, "Extended_MAUJDSServicePool"));
                extended.addProperty("MAUJDSDespatchType", getString(row, headerIndex, "Extended_MAUJDSDespatchType"));
                extended.addProperty("MAUJDSShipmentType", getString(row, headerIndex, "Extended_MAUJDSShipmentType"));
                extended.addProperty("MAUJDSPODType", getString(row, headerIndex, "Extended_MAUJDSPODType"));
                extended.addProperty("MAUJDSDeliveryInstructions1", getString(row, headerIndex, "Extended_MAUJDSDeliveryInstructions1"));
                extended.addProperty("MAUJDSDeliveryInstructions2", getString(row, headerIndex, "Extended_MAUJDSDeliveryInstructions2"));
                extended.addProperty("MAUJDSBranchNumber", getString(row, headerIndex, "Extended_MAUJDSBranchNumber"));
                extended.addProperty("MAUJDSSubOrderType", getString(row, headerIndex, "MAUJDSSubOrderType"));
                // If API expects explicit nulls, use JsonNull.INSTANCE (else omit):
                // extended.add("MAUJDSPaymentStatus", JsonNull.INSTANCE);
                // extended.add("MAUJDSPaymentType1", JsonNull.INSTANCE);
                // extended.add("MAUJDSPaymentType2", JsonNull.INSTANCE);
                // extended.add("MAUJDSPaymentType3", JsonNull.INSTANCE);

                order.add("Extended", extended);

                // Other fields
                order.addProperty("IncotermId", getString(row, headerIndex, "IncotermId"));
                order.addProperty("OrderPlacedDateTime", getString(row, headerIndex, "OrderPlacedDateTime"));
                order.addProperty("OrderType", getString(row, headerIndex, "OrderType"));
                order.addProperty("OriginFacilityId", getString(row, headerIndex, "OriginFacilityId"));

                // Contacts (dynamic: supports _0_, _1_, _2_, ...)
                JsonArray contacts = buildContactsArray(row, headerIndex);
                if (contacts.size() > 0) order.add("OriginalOrderContact", contacts);

                // OriginalOrderId
                String orderId = getString(row, headerIndex, "OriginalOrderId");
                order.addProperty("OriginalOrderId", orderId);

                // Lines (dynamic: supports _1_, _2_, ...; also works if you use _0_ start)
                BuildLinesResult linesResult = buildOrderLinesDynamic(row, headerIndex);
                if (linesResult.lines.size() == 0) {
                    // No valid lines → mark failed and skip this row from payload
                    markResult(row, headerIndex, "Failed - " + (linesResult.errors.isEmpty()
                            ? "No valid OriginalOrderLine groups"
                            : String.join("; ", linesResult.errors)));
                    continue;
                }
                order.add("OriginalOrderLine", linesResult.lines);

                // Pickup & PreShip
                order.addProperty("PickupStartDateTime", getString(row, headerIndex, "PickupStartDateTime"));
                order.addProperty("PreShipConfirmRequired", getBoolean(row, headerIndex, "PreShipConfirmRequired", true));

                // Add to bulk
                dataArray.add(order);
                orderIdToRowMap.put(orderId.trim(), row);
            }

            // Root wrapper
            JsonObject rootWrapper = new JsonObject();
            rootWrapper.add("Data", dataArray);
            String jsonBody = rootWrapper.toString();

            // --- POST ---
            String response = callPOST(jsonBody, token);
            System.out.println("Request Body:\n" + jsonBody);
            System.out.println("Response:\n" + response);

            // --- Parse response & mark results ---
//            ApiOutcome outcome = parseApiOutcome(response);
//
//            for (Map.Entry<String, Row> entry : orderIdToRowMap.entrySet()) {
//                String oid = entry.getKey();
//                Row row = entry.getValue();
//
//                String cellVal;
//                if (dataArray.size() == 0) {
//                    cellVal = "Failed - Payload empty";
//                } else if (outcome.failedOrderIds.contains(oid)) {
//                    cellVal = "Failed - " + (outcome.summary.isEmpty() ? "See response" : outcome.summary);
//                } else if (outcome.success && outcome.successCount > 0) {
//                    cellVal = "Passed";
//                } else {
//                    cellVal = "Failed - " + (outcome.summary.isEmpty() ? "No error message" : outcome.summary);
//                }
//                markResult(row, headerIndex, cellVal);
//            }

                        // Parse response
            JsonObject respJson = JsonParser.parseString(response).getAsJsonObject();
            Set<String> failedOrderIds = new HashSet<>();

            if (respJson.has("data")) {
                JsonObject data = respJson.getAsJsonObject("data");
                if (data.has("FailedRecords") && data.get("FailedRecords").isJsonArray()) {
                    JsonArray failedRecords = data.getAsJsonArray("FailedRecords");
                    for (JsonElement elem : failedRecords) {
                        JsonObject failed = elem.getAsJsonObject();
                        if (failed.has("OriginalOrderId")) {
                            String failedId = failed.get("OriginalOrderId").getAsString().trim();
                            failedOrderIds.add(failedId);
                        }
                    }
                }
            }

            // Mark each row
            for (Map.Entry<String, Row> entry : orderIdToRowMap.entrySet()) {
                String orderId = entry.getKey();
                Row row = entry.getValue();
                Cell resultCell = row.getCell(1); // Column B
                if (resultCell == null) resultCell = row.createCell(1, CellType.STRING);
                resultCell.setCellValue(failedOrderIds.contains(orderId) ? "Failed" : "Passed");
            }








            // Save workbook
            if (fis != null) fis.close();
            try (FileOutputStream fos = new FileOutputStream(filePath)) {
                workbook.write(fos);
            }
            workbook.close();

        } catch (Exception e) {
            System.err.println("❌ Error in execution: " + e.getMessage());
            e.printStackTrace();
        } finally {
            if (fis != null) try { fis.close(); } catch (IOException ignore) {}
            if (workbook != null) try { workbook.close(); } catch (IOException ignore) {}
        }
    }

    // ========= Token & POST =========

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

    public static String callPOST(String body, String token) throws IOException {
        ExcelReader reader = new ExcelReader();
        String BASE_URL             = reader.getCellValueByHeader(1, "BASE_URL");
        String SelectedOrganization = reader.getCellValueByHeader(1, "SelectedOrganization");
        String SelectedLocation     = reader.getCellValueByHeader(1, "SelectedLocation");
        reader.close();

        OkHttpClient client = new OkHttpClient();
        MediaType mediaType = MediaType.parse("application/json");
        RequestBody requestBody = RequestBody.create(mediaType, body);

        Request request = new Request.Builder()
                .url(BASE_URL + "/dcorder/api/dcorder/originalOrder/bulkImport")
                .method("POST", requestBody)
                .addHeader("Content-Type", "application/json")
                .addHeader("SelectedOrganization", SelectedOrganization)
                .addHeader("SelectedLocation", SelectedLocation)
                .addHeader("Authorization", "Bearer " + token)
                .build();

        Response response = client.newCall(request).execute();
        return response.body() != null ? response.body().string() : null;
    }

    // ========= Helpers: headers, cells =========

    private static Map<String, Integer> buildHeaderIndexMap(Sheet sheet) {
        Map<String, Integer> map = new LinkedHashMap<>();
        Row header = sheet.getRow(0);
        if (header == null) return map;
        for (int c = 0; c < header.getLastCellNum(); c++) {
            Cell cell = header.getCell(c);
            if (cell == null) continue;
            String name = cell.getStringCellValue();
            if (name != null) map.put(name.trim(), c);
        }
        return map;
    }

    private static String getString(Row row, Map<String, Integer> headerIndex, String colName) {
        Integer idx = headerIndex.get(colName);
        if (idx == null) return "";
        Cell cell = row.getCell(idx);
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING: return cell.getStringCellValue().trim();
            case NUMERIC:
                double d = cell.getNumericCellValue();
                return (d == Math.rint(d)) ? String.valueOf((long) d) : String.valueOf(d);
            case BOOLEAN: return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try { return cell.getStringCellValue().trim(); }
                catch (IllegalStateException e) {
                    double dv = cell.getNumericCellValue();
                    return (dv == Math.rint(dv)) ? String.valueOf((long) dv) : String.valueOf(dv);
                }
            default: return "";
        }
    }

    private static int getInt(Row row, Map<String, Integer> headerIndex, String colName) {
        Integer idx = headerIndex.get(colName);
        if (idx == null) return 0;
        Cell cell = row.getCell(idx);
        if (cell == null) return 0;
        switch (cell.getCellType()) {
            case NUMERIC: return (int) Math.round(cell.getNumericCellValue());
            case STRING:
                try { return Integer.parseInt(cell.getStringCellValue().trim()); }
                catch (NumberFormatException e) { return 0; }
            case FORMULA:
                try { return (int) Math.round(cell.getNumericCellValue()); }
                catch (Exception e) { return 0; }
            default: return 0;
        }
    }

    private static boolean getBoolean(Row row, Map<String, Integer> headerIndex, String colName, boolean defaultVal) {
        String s = getString(row, headerIndex, colName);
        if (s.isEmpty()) return defaultVal;
        s = s.trim().toLowerCase(Locale.ROOT);
        return s.equals("y") || s.equals("yes") || s.equals("true") || s.equals("1");
    }

    // ========= Destination & Bill-To =========

    /** Normalize text copied from Excel: trim, strip enclosing quotes (even multiple), collapse {{...}} to {...}, fix doubled quotes. */
    private static String normalizeCellText(String s) {
        if (s == null) return null;
        String t = s.trim();

        // Strip enclosing quotes repeatedly (covers ""{{...}}"" or '"... "')
        while ((t.startsWith("\"") && t.endsWith("\"")) || (t.startsWith("'") && t.endsWith("'"))) {
            t = t.substring(1, t.length() - 1).trim();
        }

        // Collapse double curly braces => {{...}} -> {...}
        if (t.startsWith("{{") && t.endsWith("}}")) {
            t = t.substring(1, t.length() - 1).trim();
        }

        // Fix doubled quotes inside JSON: ""Address1"" -> "Address1"
        t = t.replace("\"\"", "\"");

        return t;
    }

    /** Quick pre-check to decide if the string looks like a JSON object. */
    private static boolean looksLikeJsonObject(String t) {
        if (t == null) return false;
        t = t.trim();
        return t.startsWith("{") && t.endsWith("}");
    }

    /** Remove JS-style comments from JSON before parsing. */
    private static String stripJsonComments(String s) {
        if (s == null) return null;
        String noSingle = s.replaceAll("//.*", "");       // remove // comments
        return noSingle.replaceAll("/\\*.*?\\*/", "");    // remove /* */ comments
    }

    /** Try to parse a normalized JSON object; returns null if parsing fails. */
    private static JsonObject tryParseJsonObject(String normalized) {
        if (normalized == null) return null;
        String s = stripJsonComments(normalized).trim();
        if (!looksLikeJsonObject(s)) return null;
        try {
            return JsonParser.parseString(s).getAsJsonObject();
        } catch (Exception ex) {
            return null;
        }
    }

    /** Extract a facility code from a messy string by removing non [A-Za-z0-9_-] chars. */
    private static String extractFacilityCode(String s) {
        if (s == null) return null;
        String cleaned = s.trim();

        // Drop any surrounding quotes repeatedly
        while ((cleaned.startsWith("\"") && cleaned.endsWith("\"")) || (cleaned.startsWith("'") && cleaned.endsWith("'"))) {
            cleaned = cleaned.substring(1, cleaned.length() - 1).trim();
        }

        // Remove everything except letters/digits/_/-
        cleaned = cleaned.replaceAll("[^A-Za-z0-9_\\-]", "");

        // Basic sanity: non-empty and reasonable length
        if (cleaned.matches("^[A-Za-z0-9_\\-]{4,}$")) {
            return cleaned;
        }
        return null;
    }

    // Recognize DestinationAddress as JSON or facility code; also accept alias and prefixed address columns
    private static void addDestination(Row row, Map<String, Integer> headerIndex, JsonObject order) {
        // 1) Read the DestinationAddress cell and normalize it
        String raw = getString(row, headerIndex, "DestinationAddress");
        String normalized = normalizeCellText(raw);

        if (normalized != null && !normalized.isEmpty()) {
            // Try JSON address first
            JsonObject jsonAddr = tryParseJsonObject(normalized);
            if (jsonAddr != null) {
                order.add("DestinationAddress", jsonAddr);
                return;
            }

            // Else treat as facility code (robust extraction)
            String fac = extractFacilityCode(normalized);
            if (fac != null) {
                order.addProperty("DestinationFacilityId", fac);
                return;
            }
            // If neither worked, continue to other sources below
        }

        // 2) Dedicated DestinationFacilityId column (if present)
        String destFacilityId = normalizeCellText(getString(row, headerIndex, "DestinationFacilityId"));
        if (destFacilityId != null && destFacilityId.matches("^[A-Za-z0-9_\\-]{4,}$")) {
            order.addProperty("DestinationFacilityId", destFacilityId);
            return;
        }

        // 3) Optional: DestinationFacilityAlias column
        String destFacilityAlias = normalizeCellText(getString(row, headerIndex, "DestinationFacilityAlias"));
        if (destFacilityAlias != null && !destFacilityAlias.isEmpty()) {
            order.addProperty("DestinationFacilityAlias", destFacilityAlias);
            // Do not return; we may also assemble address from prefix if present
        }

        // 4) Assemble address from prefixed columns DestinationAddress_*
        JsonObject destFromPrefix = buildObjectFromPrefixedColumns(row, headerIndex, "DestinationAddress_");
        if (destFromPrefix != null) {
            order.add("DestinationAddress", destFromPrefix);
        }
    }

    // BillToAddress: JSON or prefixed columns
    private static void addBillTo(Row row, Map<String, Integer> headerIndex, JsonObject order) {
        String raw = getString(row, headerIndex, "BillToAddress");
        String normalized = normalizeCellText(raw);
        if (normalized != null && !normalized.isEmpty()) {
            if (looksLikeJsonObject(normalized)) {
                JsonObject bill = tryParseJsonObject(normalized);
                if (bill != null) {
                    order.add("BillToAddress", bill);
                    return;
                } else {
                    System.err.println("⚠️ BillToAddress looks like JSON but parsing failed.");
                }
            }
        }
        // Assemble from BillToAddress_* prefixed columns
        JsonObject billFromPrefix = buildObjectFromPrefixedColumns(row, headerIndex, "BillToAddress_");
        if (billFromPrefix != null) {
            order.add("BillToAddress", billFromPrefix);
        }
    }

    private static JsonObject buildObjectFromPrefixedColumns(Row row, Map<String, Integer> headerIndex, String prefix) {
        JsonObject obj = new JsonObject();
        boolean any = false;
        for (Map.Entry<String, Integer> entry : headerIndex.entrySet()) {
            String h = entry.getKey();
            if (h.startsWith(prefix)) {
                String field = h.substring(prefix.length());
                String val = getString(row, headerIndex, h);
                if (val != null && !val.trim().isEmpty()) {
                    obj.addProperty(field, val.trim());
                    any = true;
                }
            }
        }
        return any ? obj : null;
    }

    /**
     * Keep only one destination type; prefer Address > FacilityId > Alias.
     * (Avoid business validation conflicts and DCO::750 when nothing is present.)
     */
    private static void validateDestination(JsonObject order) {
        boolean hasAddr  = order.has("DestinationAddress");
        boolean hasFac   = order.has("DestinationFacilityId");
        boolean hasAlias = order.has("DestinationFacilityAlias");

        int count = (hasAddr ? 1 : 0) + (hasFac ? 1 : 0) + (hasAlias ? 1 : 0);
        if (count > 1) {
            System.err.println("⚠️ Multiple destination fields present; prefer Address > FacilityId > Alias.");
            if (hasAddr) {
                if (hasFac)   order.remove("DestinationFacilityId");
                if (hasAlias) order.remove("DestinationFacilityAlias");
            } else if (hasFac && hasAlias) {
                order.remove("DestinationFacilityAlias");
            }
        }
        if (count == 0) {
            System.err.println("⚠️ Neither DestinationAddress nor DestinationFacilityId supplied.");
        }
    }

    // ========= Contacts & Lines (dynamic) =========

    private static JsonArray buildContactsArray(Row row, Map<String, Integer> headerIndex) {
        // Supports either 0-based or 1-based index
        Pattern typePattern = Pattern.compile("^OriginalOrderContact_(\\d+)_ContactTypeId$");
        Set<Integer> indices = new TreeSet<>();
        for (String h : headerIndex.keySet()) {
            Matcher m = typePattern.matcher(h);
            if (m.find()) indices.add(Integer.parseInt(m.group(1)));
        }
        JsonArray contacts = new JsonArray();
        for (Integer i : indices) {
            String type = getString(row, headerIndex, "OriginalOrderContact_" + i + "_ContactTypeId");
            String fn   = getString(row, headerIndex, "OriginalOrderContact_" + i + "_Contact_FirstName");
            String ln   = getString(row, headerIndex, "OriginalOrderContact_" + i + "_Contact_LastName");
            if (type.isEmpty() && fn.isEmpty() && ln.isEmpty()) continue;

            JsonObject wrapper = new JsonObject();
            wrapper.addProperty("ContactTypeId", type);

            JsonObject contact = new JsonObject();
            contact.addProperty("FirstName", fn);
            contact.addProperty("LastName", ln);

            wrapper.add("Contact", contact);
            contacts.add(wrapper);
        }
        return contacts;
    }

    private static class BuildLinesResult {
        JsonArray lines = new JsonArray();
        List<String> errors = new ArrayList<>();
        List<String> warnings = new ArrayList<>();
    }

    private static BuildLinesResult buildOrderLinesDynamic(Row row, Map<String, Integer> headerIndex) {
        BuildLinesResult result = new BuildLinesResult();

        // Discover indices n by checking for ItemId columns
        Pattern itemField = Pattern.compile("^OriginalOrderLine_(\\d+)_ItemId$");
        Set<Integer> lineIndices = new TreeSet<>();
        for (String h : headerIndex.keySet()) {
            Matcher m = itemField.matcher(h);
            if (m.find()) lineIndices.add(Integer.parseInt(m.group(1)));
        }

        for (Integer n : lineIndices) {
            String itemId   = getString(row, headerIndex, "OriginalOrderLine_" + n + "_ItemId");
            int qty         = getInt(row, headerIndex, "OriginalOrderLine_" + n + "_OrderedQuantity");
            String lineId   = getString(row, headerIndex, "OriginalOrderLine_" + n + "_OriginalOrderLineId");
            String uom      = getString(row, headerIndex, "OriginalOrderLine_" + n + "_QuantityUomId");
            int unitPrice   = getInt(row, headerIndex, "OriginalOrderLine_" + n + "_UnitPrice");

            boolean anyPresent = (itemId != null && !itemId.isEmpty())
                    || (qty != 0) || (lineId != null && !lineId.isEmpty())
                    || (uom != null && !uom.isEmpty()) || (unitPrice != 0);

            if (!anyPresent) continue; // nothing in this group

            List<String> errs = new ArrayList<>();
            if (itemId == null || itemId.isEmpty()) errs.add("ItemId is required");
            if (lineId == null || lineId.isEmpty()) errs.add("OriginalOrderLineId is required");
            if (uom == null || uom.isEmpty()) errs.add("QuantityUomId is required");
            if (qty <= 0) errs.add("OrderedQuantity must be > 0");

            if (!errs.isEmpty()) {
                result.warnings.add("Line " + n + " ignored (" + String.join(", ", errs) + ")");
                continue;
            }

            JsonObject line = new JsonObject();
            line.addProperty("ItemId", itemId);
            line.addProperty("OrderedQuantity", qty);
            line.addProperty("OriginalOrderLineId", lineId);
            line.addProperty("QuantityUomId", uom);
            line.addProperty("UnitPrice", unitPrice);

            result.lines.add(line);
        }

        if (result.lines.size() == 0) {
            result.errors.add("At least one valid OriginalOrderLine is required.");
        }
        for (String w : result.warnings) {
            System.out.println("⚠️ " + w);
        }
        return result;
    }

    // ========= Response parsing & marking =========

    static class ApiOutcome {
        boolean success = false;
        int successCount = 0;
        int failedCount = 0;
        Set<String> failedOrderIds = new HashSet<>();
        String summary = "";
    }

    private static ApiOutcome parseApiOutcome(String responseBody) {
        ApiOutcome outcome = new ApiOutcome();

        if (responseBody == null || responseBody.trim().isEmpty()) {
            outcome.summary = "Empty response body";
            return outcome;
        }

        try {
            JsonObject resp = JsonParser.parseString(responseBody).getAsJsonObject();

            if (resp.has("success") && !resp.get("success").isJsonNull()) {
                outcome.success = resp.get("success").getAsBoolean();
            }

            if (resp.has("data") && resp.get("data").isJsonObject()) {
                JsonObject data = resp.getAsJsonObject("data");
                if (data.has("SuccessCount")) {
                    try { outcome.successCount = data.get("SuccessCount").getAsInt(); } catch (Exception ignore) {}
                }
                if (data.has("FailedCount")) {
                    try { outcome.failedCount = data.get("FailedCount").getAsInt(); } catch (Exception ignore) {}
                }
                if (data.has("FailedRecords") && data.get("FailedRecords").isJsonArray()) {
                    JsonArray failedRecords = data.getAsJsonArray("FailedRecords");
                    for (JsonElement el : failedRecords) {
                        JsonObject rec = el.getAsJsonObject();
                        if (rec.has("OriginalOrderId")) {
                            outcome.failedOrderIds.add(rec.get("OriginalOrderId").getAsString().trim());
                        }
                    }
                }
            }

            // Messages → ShortDescription/Description + Code
            StringBuilder sb = new StringBuilder();
            if (resp.has("messages") && resp.get("messages").isJsonObject()) {
                JsonObject messages = resp.getAsJsonObject("messages");
                if (messages.has("Message") && messages.get("Message").isJsonArray()) {
                    JsonArray arr = messages.getAsJsonArray("Message");
                    for (JsonElement el : arr) {
                        JsonObject msg = el.getAsJsonObject();
                        String type = msg.has("Type") ? msg.get("Type").getAsString() : "";
                        if ("ERROR".equalsIgnoreCase(type)) {
                            String code = msg.has("Code") && !msg.get("Code").isJsonNull() ? msg.get("Code").getAsString() : "";
                            String shortDesc = msg.has("ShortDescription") && !msg.get("ShortDescription").isJsonNull()
                                    ? msg.get("ShortDescription").getAsString() : "";
                            String desc = msg.has("Description") && !msg.get("Description").isJsonNull()
                                    ? msg.get("Description").getAsString() : "";

                            if (sb.length() > 0) sb.append("; ");
                            if (!shortDesc.isEmpty()) sb.append(shortDesc);
                            else if (!desc.isEmpty()) sb.append(desc);
                            if (!code.isEmpty()) sb.append(" (").append(code).append(")");
                        }
                    }
                }
            }

            // Exceptions (fallback)
            if (sb.length() == 0 && resp.has("exceptions") && resp.get("exceptions").isJsonArray()) {
                JsonArray exArr = resp.getAsJsonArray("exceptions");
                for (JsonElement el : exArr) {
                    JsonObject ex = el.getAsJsonObject();
                    String key = ex.has("messageKey") && !ex.get("messageKey").isJsonNull()
                            ? ex.get("messageKey").getAsString() : "";
                    String msg = ex.has("message") && !ex.get("message").isJsonNull()
                            ? ex.get("message").getAsString() : "";
                    if (!msg.isEmpty()) {
                        if (sb.length() > 0) sb.append("; ");
                        sb.append(msg);
                        if (!key.isEmpty()) sb.append(" (").append(key).append(")");
                    }
                }
            }

            outcome.summary = sb.toString();
            if (outcome.summary.isEmpty() && outcome.success) {
                outcome.summary = "Success";
            } else if (outcome.summary.isEmpty()) {
                outcome.summary = "No error message provided";
            }

        } catch (Exception e) {
            outcome.summary = "Response parse error: " + e.getMessage();
        }

        return outcome;
    }

    private static void markResult(Row row, Map<String, Integer> headerIndex, String value) {
        Integer resultColIdx = headerIndex.get("Result");
        if (resultColIdx == null) return;
        Cell resultCell = row.getCell(resultColIdx);
        if (resultCell == null) resultCell = row.createCell(resultColIdx, CellType.STRING);
        resultCell.setCellValue(value);
    }

    // ========= Utility =========

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
