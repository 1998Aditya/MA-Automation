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
//public class Main5_EcomOrderCreation {
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
//            Sheet sheet = workbook.getSheet("ECOM Order");
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
//                order.addProperty("Cancelled", getString(row, 2));
//                order.addProperty("DeliveryEndDateTime", getString(row, 3));
//                order.addProperty("DeliveryStartDateTime", getString(row, 4));
//
//
//                String destinationRaw = getString(row, 5);
//                JsonObject destinationJson = JsonParser.parseString("{" + destinationRaw + "}").getAsJsonObject();
//                order.add("DestinationAddress", destinationJson);
//
//                String BillToAddressRaw = getString(row, 6);
//                JsonObject BillToAddressJson = JsonParser.parseString("{" + BillToAddressRaw + "}").getAsJsonObject();
//                order.add("BillToAddress", BillToAddressJson);
//
//
//                JsonObject extended = new JsonObject();
//                extended.addProperty("MAUJDSDeliveryType", getString(row, 7));
//                extended.addProperty("MAUJDSPrepaid", getString(row, 8));
//                extended.addProperty("MAUJDSPaymentType1Value", getString(row, 9));
//                extended.addProperty("MAUJDSPaymentType2Value", getString(row, 10));
//                extended.addProperty("MAUJDSPaymentType3Value", getString(row, 11));
//                extended.addProperty("MAUJDSPostageCharge", getInt(row, 12));
//                extended.addProperty("MAUJDSSignatory", "");
//                extended.add("MAUJDSPaymentStatus", null);
//                extended.add("MAUJDSPaymentType1", null);
//                extended.add("MAUJDSPaymentType2", null);
//                extended.add("MAUJDSPaymentType3", null);
//                extended.addProperty("MAUJDSFasciaCode", getString(row, 18));
//                extended.addProperty("MAUJDSFasciaDescription", getString(row, 19));
//                extended.addProperty("MAUJDSWebOrderNumber", getString(row, 20));
//                extended.addProperty("MAUJDSTotalPrice", getInt(row, 21));
//                extended.addProperty("MAUJDSPODType", getString(row, 22));
//                extended.addProperty("MAUJDSServicePool", getString(row, 23));
//                extended.addProperty("MAUJDSShipmentType", getString(row, 24));
//                extended.addProperty("MAUJDSDespatchType", getString(row, 25));
//                extended.addProperty("MAUJDSDeliveryInstructions1", "");
//                extended.addProperty("MAUJDSDeliveryInstructions2", "");
//                extended.addProperty("MAUJDSBranchNumber", getString(row, 28));
//                extended.addProperty("MAUJDSSubOrderType", getString(row, 29));
//
//
//
//                order.add("Extended", extended);
//                order.addProperty("IncotermId", getString(row, 30));
//                order.addProperty("OrderPlacedDateTime", getString(row, 31));
//                order.addProperty("OrderType", getString(row, 32));
//                order.addProperty("OriginFacilityId", getString(row, 33));
//
//                JsonArray contacts = new JsonArray();
//                JsonObject contactWrapper = new JsonObject();
//                contactWrapper.addProperty("ContactTypeId", getString(row, 34));
//
//                JsonObject contact = new JsonObject();
//                contact.addProperty("FirstName", getString(row, 35));
//                contact.addProperty("LastName", getString(row, 36));
//
//                contactWrapper.add("Contact", contact);
//                contacts.add(contactWrapper);
//                order.add("OriginalOrderContact", contacts);
//
//                String orderId = getString(row, 37);
//                order.addProperty("OriginalOrderId", orderId);
//
//                JsonArray orderLines = new JsonArray();
//                JsonObject line = new JsonObject();
//                line.addProperty("ItemId", getString(row, 38));
//                line.addProperty("OrderedQuantity", getInt(row, 39));
//                line.addProperty("OriginalOrderLineId", getString(row, 40));
//                line.addProperty("QuantityUomId", getString(row, 41));
//                line.addProperty("UnitPrice", getInt(row, 42));
//                orderLines.add(line);
//                order.add("OriginalOrderLine", orderLines);
//
//                order.addProperty("PickupStartDateTime", getString(row, 43));
//                order.addProperty("PreShipConfirmRequired",getString(row, 44));
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

public class Main5_EcomOrderCreation {

    // ===== Simple wrapper to carry HTTP code + body =====
    static class HttpResult {
        int statusCode;
        String body;
        boolean isSuccessful;
    }

    // ===== Entry point =====
    public static void main(String testcaseToRun, String filePath) throws IOException {
        closeExcelIfOpen();
        FileInputStream fis = null;
        Workbook workbook = null;

        try {
            String token = getAuthTokenFromExcel();
            if (token == null) {
                System.err.println("Failed to retrieve access token.");
                return;
            }

            fis = new FileInputStream(filePath);
            workbook = new XSSFWorkbook(fis);

            Sheet sheet = workbook.getSheet("ECOM Order");
            if (sheet == null) throw new IllegalStateException("Sheet 'ECOM Order' not found in: " + filePath);

            Map<String, Integer> headerIndex = buildHeaderIndexMap(sheet);

            JsonArray dataArray = new JsonArray();
            Map<String, Row> orderIdToRowMap = new HashMap<>();

            // Iterate rows
            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;

                String testcase = getString(row, headerIndex, "Testcase");
                if (!testcase.equalsIgnoreCase(testcaseToRun)) continue;

                JsonObject order = new JsonObject();

                // ---- Scalars ----
                // NOTE: Keep Cancelled as String to match your original payload style
                order.addProperty("Cancelled", getString(row, headerIndex, "Cancelled"));
                order.addProperty("DeliveryEndDateTime", getString(row, headerIndex, "DeliveryEndDateTime"));
                order.addProperty("DeliveryStartDateTime", getString(row, headerIndex, "DeliveryStartDateTime"));

                // ---- Destination & BillTo ----
                addDestination(row, headerIndex, order);   // Accepts JSON or facility id
                addBillTo(row, headerIndex, order);        // Accepts JSON or prefixed columns

                // ---- Extended block ----
                JsonObject extended = new JsonObject();
                extended.addProperty("MAUJDSDeliveryType", getString(row, headerIndex, "MAUJDSDeliveryType"));
                extended.addProperty("MAUJDSPrepaid", getString(row, headerIndex, "MAUJDSPrepaid"));
                extended.addProperty("MAUJDSPaymentType1Value", getString(row, headerIndex, "MAUJDSPaymentType1Value"));
                extended.addProperty("MAUJDSPaymentType2Value", getString(row, headerIndex, "MAUJDSPaymentType2Value"));
                extended.addProperty("MAUJDSPaymentType3Value", getString(row, headerIndex, "MAUJDSPaymentType3Value"));
                extended.addProperty("MAUJDSPostageCharge", getInt(row, headerIndex, "MAUJDSPostageCharge"));
                extended.addProperty("MAUJDSSignatory", getString(row, headerIndex, "MAUJDSSignatory"));
                extended.add("MAUJDSPaymentStatus", JsonNull.INSTANCE);
                extended.add("MAUJDSPaymentType1", JsonNull.INSTANCE);
                extended.add("MAUJDSPaymentType2", JsonNull.INSTANCE);
                extended.add("MAUJDSPaymentType3", JsonNull.INSTANCE);
                extended.addProperty("MAUJDSFasciaCode", getString(row, headerIndex, "MAUJDSFasciaCode"));
                extended.addProperty("MAUJDSFasciaDescription", getString(row, headerIndex, "MAUJDSFasciaDescription"));
                extended.addProperty("MAUJDSWebOrderNumber", getString(row, headerIndex, "MAUJDSWebOrderNumber"));
                extended.addProperty("MAUJDSTotalPrice", getInt(row, headerIndex, "MAUJDSTotalPrice"));
                extended.addProperty("MAUJDSPODType", getString(row, headerIndex, "MAUJDSPODType"));
                extended.addProperty("MAUJDSServicePool", getString(row, headerIndex, "MAUJDSServicePool"));
                extended.addProperty("MAUJDSShipmentType", getString(row, headerIndex, "MAUJDSShipmentType"));
                extended.addProperty("MAUJDSDespatchType", getString(row, headerIndex, "MAUJDSDespatchType"));
                extended.addProperty("MAUJDSDeliveryInstructions1", getString(row, headerIndex, "MAUJDSDeliveryInstructions1"));
                extended.addProperty("MAUJDSDeliveryInstructions2", getString(row, headerIndex, "MAUJDSDeliveryInstructions2"));
                extended.addProperty("MAUJDSBranchNumber", getString(row, headerIndex, "MAUJDSBranchNumber"));
                extended.addProperty("MAUJDSSubOrderType", getString(row, headerIndex, "MAUJDSSubOrderType"));
                order.add("Extended", extended);

                // ---- Others ----
                order.addProperty("IncotermId", getString(row, headerIndex, "IncotermId"));
                order.addProperty("OrderPlacedDateTime", getString(row, headerIndex, "OrderPlacedDateTime"));
                order.addProperty("OrderType", getString(row, headerIndex, "OrderType"));
                order.addProperty("OriginFacilityId", getString(row, headerIndex, "OriginFacilityId"));

                // ---- Contacts (optional) ----
                JsonArray contacts = buildContactsArray(row, headerIndex);
                if (contacts.size() > 0) order.add("OriginalOrderContact", contacts);

                // ---- Order Id ----
                String orderId = getString(row, headerIndex, "OriginalOrderId");
                order.addProperty("OriginalOrderId", orderId);

                // ---- Dynamic order lines for any indices {n} present (1..N) ----
                BuildLinesResult linesResult = buildOrderLinesDynamic(row, headerIndex);
                if (linesResult.lines.size() == 0) {
                    // No valid lines: mark and skip this row
                    markResult(row, headerIndex, "Failed - " + String.join("; ",
                            linesResult.errors.isEmpty()
                                    ? Collections.singletonList("No valid OriginalOrderLine groups")
                                    : linesResult.errors));
                    continue;
                }
                order.add("OriginalOrderLine", linesResult.lines);

                // ---- Pickup & PreShip ----
                order.addProperty("PickupStartDateTime", getString(row, headerIndex, "PickupStartDateTime"));
                // Keep as string to match your original, but if the API expects boolean, you can switch:
                order.addProperty("PreShipConfirmRequired", getString(row, headerIndex, "PreShipConfirmRequired"));

                // Accumulate
                dataArray.add(order);
                orderIdToRowMap.put(orderId != null ? orderId.trim() : "", row);
            }

            // Wrap and send
            JsonObject rootWrapper = new JsonObject();
            rootWrapper.add("Data", dataArray);
            String jsonBody = rootWrapper.toString();

            HttpResult http = callPOST(jsonBody, token);
            System.out.println("Request Body:\n" + jsonBody);
            System.out.println("HTTP Status: " + http.statusCode + " (successful=" + http.isSuccessful + ")");
            System.out.println("Response:\n" + http.body);

            // Parse API response
            ApiOutcome outcome = parseApiOutcome(http);

            // Mark Result cell by orderId
            for (Map.Entry<String, Row> entry : orderIdToRowMap.entrySet()) {
                String oid = entry.getKey();
                Row row = entry.getValue();

                String cellVal;
                if (outcome.failedOrderIds.contains(oid)) {
                    cellVal = "Failed (HTTP " + http.statusCode + ") - " + outcome.summary;
                } else if (dataArray.size() == 0) {
                    cellVal = "Failed - Payload empty";
                } else {
                    if (outcome.success && outcome.successCount > 0) {
                        cellVal = "Passed (HTTP " + http.statusCode + ")";
                    } else {
                        cellVal = "Failed (HTTP " + http.statusCode + ") - " + outcome.summary;
                    }
                }

                // Prefer "Result" by header; else fallback to column B (index 1)
                if (headerIndex.containsKey("Result")) {
                    markResult(row, headerIndex, cellVal);
                } else {
                    Cell resultCell = row.getCell(1);
                    if (resultCell == null) resultCell = row.createCell(1, CellType.STRING);
                    resultCell.setCellValue(cellVal);
                }
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

    // ===== Token & POST =====
    public static String getAuthTokenFromExcel() throws IOException {
        ExcelReader reader = new ExcelReader();
        String LOGIN_URL = reader.getCellValueByHeader(1, "LOGIN_URL");
        String UIUsername = reader.getCellValueByHeader(1, "username");
        String UIPassword = reader.getCellValueByHeader(1, "password");
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

    // Return HttpResult (status + body); also add a String-returning overload if you need compatibility.
    public static HttpResult callPOST(String body, String token) throws IOException {
        ExcelReader reader = new ExcelReader();
        String BASE_URL = reader.getCellValueByHeader(1, "BASE_URL");
        String SelectedOrganization = reader.getCellValueByHeader(1, "SelectedOrganization");
        String SelectedLocation = reader.getCellValueByHeader(1, "SelectedLocation");
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

        HttpResult result = new HttpResult();
        result.statusCode = response.code();
        result.isSuccessful = response.isSuccessful();
        result.body = (response.body() != null) ? response.body().string() : "";
        return result;
    }

    // ===== Helpers: headers, cell reads, destination/billTo, contacts =====
    private static Map<String, Integer> buildHeaderIndexMap(Sheet sheet) {
        Map<String, Integer> map = new HashMap<>();
        Row header = sheet.getRow(0);
        if (header == null) return map;
        for (int c = 0; c < header.getLastCellNum(); c++) {
            Cell cell = header.getCell(c);
            if (cell == null) continue;
            if (cell.getCellType() == CellType.STRING) {
                String name = cell.getStringCellValue();
                if (name != null) map.put(name.trim(), c);
            } else {
                String name = new DataFormatter().formatCellValue(cell);
                if (name != null && !name.trim().isEmpty()) {
                    map.put(name.trim(), c);
                }
            }
        }
        return map;
    }

    private static String getString(Row row, Map<String, Integer> headerIndex, String colName) {
        Integer idx = headerIndex.get(colName);
        if (idx == null) return "";
        Cell cell = row.getCell(idx);
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                double d = cell.getNumericCellValue();
                if (d == Math.rint(d)) return String.valueOf((long) d);
                return String.valueOf(d);
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    return cell.getStringCellValue().trim();
                } catch (IllegalStateException e) {
                    double dv = cell.getNumericCellValue();
                    if (dv == Math.rint(dv)) return String.valueOf((long) dv);
                    return String.valueOf(dv);
                }
            default:
                return "";
        }
    }

    private static int getInt(Row row, Map<String, Integer> headerIndex, String colName) {
        Integer idx = headerIndex.get(colName);
        if (idx == null) return 0;
        Cell cell = row.getCell(idx);
        if (cell == null) return 0;
        switch (cell.getCellType()) {
            case NUMERIC:
                return (int) Math.round(cell.getNumericCellValue());
            case STRING:
                try {
                    return Integer.parseInt(cell.getStringCellValue().trim());
                } catch (NumberFormatException e) {
                    return 0;
                }
            case FORMULA:
                try {
                    return (int) Math.round(cell.getNumericCellValue());
                } catch (Exception e) {
                    return 0;
                }
            default:
                return 0;
        }
    }

    // DestinationAddress: JSON object or facility id (or fallback column)
    private static void addDestination(Row row, Map<String, Integer> headerIndex, JsonObject order) {
        String raw = getString(row, headerIndex, "DestinationAddress");
        if (raw != null && !raw.isEmpty()) {
            String t = raw.trim();
            // Try parse JSON
            JsonObject obj = parseObjectOrEmpty(t);
            if (obj != null && obj.size() > 0) {
                order.add("DestinationAddress", obj);
                return;
            }
            // Not JSON → treat as facility code (e.g., SHG01280)
            if (!t.isEmpty()) {
                order.addProperty("DestinationFacilityId", t);
                return;
            }
        }
        // Fallback column
        String destFacilityId = getString(row, headerIndex, "DestinationFacilityId");
        if (!destFacilityId.isEmpty()) {
            order.addProperty("DestinationFacilityId", destFacilityId);
        }
    }

    // BillToAddress: JSON or prefixed columns (e.g., BillToAddress_Address1, _City, _State, ...)
    private static void addBillTo(Row row, Map<String, Integer> headerIndex, JsonObject order) {
        String raw = getString(row, headerIndex, "BillToAddress");
        if (raw != null && !raw.isEmpty()) {
            JsonObject obj = parseObjectOrEmpty(raw.trim());
            if (obj != null && obj.size() > 0) {
                order.add("BillToAddress", obj);
                return;
            }
        }
        // Prefixed columns assembly
        JsonObject billFromPrefix = buildObjectFromPrefixedColumns(row, headerIndex, "BillToAddress_");
        if (billFromPrefix != null) {
            order.add("BillToAddress", billFromPrefix);
        }
    }

    private static JsonObject parseObjectOrEmpty(String raw) {
        if (raw == null || raw.trim().isEmpty()) return new JsonObject();
        String s = raw.trim();
        // Ensure braces
        if (!s.startsWith("{")) s = "{" + s + "}";
        try {
            return JsonParser.parseString(s).getAsJsonObject();
        } catch (Exception ex) {
            System.err.println("⚠️ Could not parse JSON object from: " + raw + " — using {}");
            return new JsonObject();
        }
    }

    // Build object from columns like DestinationAddress_Address1, _City, _State...
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

    // Contacts: support OriginalOrderContact_n_ContactTypeId and OriginalOrderContact_n_Contact_FirstName/_LastName
    private static JsonArray buildContactsArray(Row row, Map<String, Integer> headerIndex) {
        Pattern typePattern = Pattern.compile("^OriginalOrderContact_(\\d+)_ContactTypeId$");
        Set<Integer> indices = new TreeSet<>();
        for (String h : headerIndex.keySet()) {
            Matcher m = typePattern.matcher(h);
            if (m.find()) indices.add(Integer.parseInt(m.group(1)));
        }
        JsonArray contacts = new JsonArray();
        for (Integer i : indices) {
            String type = getString(row, headerIndex, "OriginalOrderContact_" + i + "_ContactTypeId");
            String fn = getString(row, headerIndex, "OriginalOrderContact_" + i + "_Contact_FirstName");
            String ln = getString(row, headerIndex, "OriginalOrderContact_" + i + "_Contact_LastName");
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

    // ===== Dynamic order lines for any indices {n} present (1..N) =====
    private static class BuildLinesResult {
        JsonArray lines = new JsonArray();
        List<String> errors = new ArrayList<>();
        List<String> warnings = new ArrayList<>();
    }

    private static BuildLinesResult buildOrderLinesDynamic(Row row, Map<String, Integer> headerIndex) {
        BuildLinesResult result = new BuildLinesResult();

        // Discover line indices by checking for ItemId columns
        Pattern itemField = Pattern.compile("^OriginalOrderLine_(\\d+)_ItemId$");
        Set<Integer> lineIndices = new TreeSet<>();
        for (String h : headerIndex.keySet()) {
            Matcher m = itemField.matcher(h);
            if (m.find()) lineIndices.add(Integer.parseInt(m.group(1)));
        }

        for (Integer n : lineIndices) {
            String itemId  = getString(row, headerIndex, "OriginalOrderLine_" + n + "_ItemId");
            int qty        = getInt(row, headerIndex, "OriginalOrderLine_" + n + "_OrderedQuantity");
            String lineId  = getString(row, headerIndex, "OriginalOrderLine_" + n + "_OriginalOrderLineId");
            String uom     = getString(row, headerIndex, "OriginalOrderLine_" + n + "_QuantityUomId");
            int unitPrice  = getInt(row, headerIndex, "OriginalOrderLine_" + n + "_UnitPrice");

            boolean anyPresent =
                    (itemId != null && !itemId.isEmpty())
                            || qty != 0
                            || (lineId != null && !lineId.isEmpty())
                            || (uom != null && !uom.isEmpty())
                            || unitPrice != 0;

            if (!anyPresent) continue; // nothing in this group

            // Validate required fields
            List<String> errs = new ArrayList<>();
            if (itemId == null || itemId.isEmpty()) errs.add("ItemId is required");
            if (lineId == null || lineId.isEmpty()) errs.add("OriginalOrderLineId is required");
            if (uom == null || uom.isEmpty()) errs.add("QuantityUomId is required");
            if (qty <= 0) errs.add("OrderedQuantity must be > 0");

            if (!errs.isEmpty()) {
                result.warnings.add("Line " + n + " ignored (" + String.join(", ", errs) + ")");
                continue; // skip invalid line
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

    // ===== Response parsing & result summarization =====
    static class ApiOutcome {
        boolean success = false;
        int successCount = 0;
        int failedCount = 0;
        Set<String> failedOrderIds = new HashSet<>();
        String summary = ""; // short human-friendly message
    }

    private static ApiOutcome parseApiOutcome(HttpResult http) {
        ApiOutcome outcome = new ApiOutcome();

        if (http.body == null || http.body.isEmpty()) {
            outcome.summary = http.isSuccessful ? "Empty response body" : "HTTP error";
            return outcome;
        }

        try {
            JsonObject resp = JsonParser.parseString(http.body).getAsJsonObject();

            // Top-level success flag if present
            if (resp.has("success")) {
                try { outcome.success = resp.get("success").getAsBoolean(); } catch (Exception ignored) {}
            }

            // Counts from 'data'
            if (resp.has("data") && resp.get("data").isJsonObject()) {
                JsonObject data = resp.getAsJsonObject("data");
                if (data.has("SuccessCount")) outcome.successCount = safeInt(data.get("SuccessCount"));
                if (data.has("FailedCount")) outcome.failedCount = safeInt(data.get("FailedCount"));

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

            // Messages (prefer ShortDescription, then Description)
            StringBuilder sb = new StringBuilder();
            if (resp.has("messages") && resp.get("messages").isJsonObject()) {
                JsonObject messages = resp.getAsJsonObject("messages");
                if (messages.has("Message") && messages.get("Message").isJsonArray()) {
                    JsonArray arr = messages.getAsJsonArray("Message");
                    for (JsonElement el : arr) {
                        JsonObject msg = el.getAsJsonObject();
                        String type = msg.has("Type") ? asStringSafe(msg.get("Type")) : "";
                        String shortDesc = msg.has("ShortDescription") ? asStringSafe(msg.get("ShortDescription")) : "";
                        String desc = msg.has("Description") ? asStringSafe(msg.get("Description")) : "";
                        String code = msg.has("Code") ? asStringSafe(msg.get("Code")) : "";

                        if ("ERROR".equalsIgnoreCase(type)) {
                            if (!shortDesc.isEmpty()) appendMsg(sb, shortDesc, code);
                            else if (!desc.isEmpty()) appendMsg(sb, desc, code);
                        }
                    }
                }
            }

            // Fallback to exceptions array
            if (sb.length() == 0 && resp.has("exceptions") && resp.get("exceptions").isJsonArray()) {
                JsonArray exArr = resp.getAsJsonArray("exceptions");
                for (JsonElement el : exArr) {
                    JsonObject ex = el.getAsJsonObject();
                    String key = ex.has("messageKey") ? asStringSafe(ex.get("messageKey")) : "";
                    String msg = ex.has("message") ? asStringSafe(ex.get("message")) : "";
                    if (!msg.isEmpty()) appendMsg(sb, msg, key);
                }
            }

            outcome.summary = sb.length() == 0
                    ? (outcome.success ? "Success" : "No error message provided")
                    : sb.toString();

        } catch (Exception e) {
            outcome.summary = "Response parse error: " + e.getMessage();
        }
        return outcome;
    }

    private static String asStringSafe(JsonElement el) {
        try { return el.isJsonNull() ? "" : el.getAsString(); } catch (Exception e) { return ""; }
    }

    private static int safeInt(JsonElement el) {
        try { return el.getAsInt(); } catch (Exception e) { return 0; }
    }

    private static void appendMsg(StringBuilder sb, String text, String code) {
        if (sb.length() > 0) sb.append("; ");
        if (code != null && !code.isEmpty()) {
            sb.append(text).append(" (").append(code).append(")");
        } else {
            sb.append(text);
        }
    }

    private static void markResult(Row row, Map<String, Integer> headerIndex, String value) {
        Integer resultColIdx = headerIndex.get("Result");
        if (resultColIdx == null) {
            // fallback to column B
            Cell cell = row.getCell(1);
            if (cell == null) cell = row.createCell(1, CellType.STRING);
            cell.setCellValue(value);
            return;
        }
        Cell resultCell = row.getCell(resultColIdx);
        if (resultCell == null) resultCell = row.createCell(resultColIdx, CellType.STRING);
        resultCell.setCellValue(value);
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
