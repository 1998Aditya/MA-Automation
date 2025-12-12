package MA_MSG_Suite_OB;

import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;
import io.github.bonigarcia.wdm.WebDriverManager;
import okhttp3.*;
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
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import com.google.gson.*;
import okhttp3.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;




public class MainA1_OrderReadyForPacking {

        private static final OkHttpClient client = new OkHttpClient();

    public static class PackingData {
        public String OLPNs;
        public List<ItemDetail> items = new ArrayList<>();

        public static class ItemDetail {
            public String itemId;
            public double quantity;

            public ItemDetail(String itemId, double quantity) {
                this.itemId = itemId;
                this.quantity = quantity;
            }
        }
    }









//        public static class PackingData {
//            String OLPNs;
//            String item1;
//            int quantities2;
//
//            @Override
//            public String toString() {
//                return "OLPNs: " + OLPNs + ", item1: " + item1 + ", quantities2: " + quantities2;
//            }
//        }

        public static void main(String filePath,String messageType) {
            try {
                List<PackingData> packingList = readPackingData(filePath);
                System.out.println("✅ Packing Data Extracted:");
                packingList.forEach(System.out::println);

                Map<String, JsonObject> jsonPayloads = buildPayloads(packingList,messageType);

                System.out.println("\n📦 Generated JSON Payloads:");
                for (Map.Entry<String, JsonObject> entry : jsonPayloads.entrySet()) {
                    System.out.println("------------------------------------------------------");
                    System.out.println("OLPN: " + entry.getKey());
                    System.out.println(entry.getValue().toString());
                }

                String token = getAuthTokenFromExcel();
                if (token == null) {
                    System.err.println("❌ Failed to authenticate.");
                    return;
                }

                triggerAPI(jsonPayloads, token);
                validateOLPNs(packingList, token,filePath,messageType);

            } catch (Exception e) {
                System.err.println("❌ Error: " + e.getMessage());
                e.printStackTrace();
            }
        }





    public static List<PackingData> readPackingData(String path) throws IOException {
        List<PackingData> packingDataList = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(path);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheet("Tasks");
            if (sheet == null) {
                System.err.println("❌ Sheet 'Tasks' not found.");
                return packingDataList;
            }

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                PackingData data = new PackingData();
                data.OLPNs = getCellValueAsString(row.getCell(4)); // OLPNs column

                // Start from column 6 (location), then every 3 columns: location, item, quantity
                for (int col = 6; col + 3 <= row.getLastCellNum(); col += 3){
//                    String itemId = getCellValueAsString(row.getCell(col + 1));
//                    double quantity = parseDoubleSafe(row.getCell(col + 2));
                   String itemId = getCellValueAsString(row.getCell(col + 2));
                    double quantity = parseDoubleSafe(row.getCell(col + 3));
                    System.out.println("ItemId: " + itemId + ", Quantity: " + quantity);

                    if (!itemId.isEmpty() && quantity > 0) {
                        data.items.add(new PackingData.ItemDetail(itemId, quantity));
                    }
                }

                if (!data.OLPNs.isEmpty() && !data.items.isEmpty()) {
                    packingDataList.add(data);
                }
            }
        }
        return packingDataList;
    }

    public static Map<String, JsonObject> buildPayloads(List<PackingData> packingList, String messageType) {
        Map<String, List<PackingData>> groupedByOLPN = new HashMap<>();
        for (PackingData data : packingList) {
            groupedByOLPN.computeIfAbsent(data.OLPNs, k -> new ArrayList<>()).add(data);
        }

        Map<String, JsonObject> payloadMap = new LinkedHashMap<>();
        for (String olpn : groupedByOLPN.keySet()) {
            JsonObject payload = new JsonObject();
            payload.addProperty("WCSOrderId", olpn);
            payload.addProperty("MessageType", messageType);
            payload.addProperty("UniqueKey", UUID.randomUUID().toString().replace("-", ""));

            JsonArray detailsArray = new JsonArray();
            for (PackingData data : groupedByOLPN.get(olpn)) {
                for (PackingData.ItemDetail item : data.items) {
                    JsonObject detail = new JsonObject();
                    detail.addProperty("ItemId", item.itemId);
                    detail.addProperty("Quantity", item.quantity);
                    detailsArray.add(detail);
                }
            }
            payload.add("Details", detailsArray);
            payloadMap.put(olpn, payload);
        }

        return payloadMap;
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


        private static void triggerAPI(Map<String, JsonObject> payloads, String token) throws IOException {
            for (Map.Entry<String, JsonObject> entry : payloads.entrySet()) {
                String olpn = entry.getKey();
                JsonObject payload = entry.getValue();

                ExcelReader reader = new ExcelReader();
                String BASE_URL             = reader.getCellValueByHeader(1, "BASE_URL");
                String SelectedOrganization = reader.getCellValueByHeader(1, "SelectedOrganization");
                String SelectedLocation     = reader.getCellValueByHeader(1, "SelectedLocation");
                reader.close();

                RequestBody body = RequestBody.create(payload.toString(), MediaType.get("application/json"));
                Request request = new Request.Builder()
                        .url(BASE_URL +"/device-integration/api/deviceintegration/process/OrderReadyForPacking_FER_Src_EP")
                        .post(body)
                        .addHeader("Authorization", "Bearer " + token)
                        .addHeader("Content-Type", "application/json")
                        .addHeader("SelectedOrganization", SelectedOrganization)
                        .addHeader("SelectedLocation", SelectedLocation)
                        .build();

                try (Response response = client.newCall(request).execute()) {
                    if (response.isSuccessful()) {
                        System.out.println("✅ Successfully sent OrderReadyForPacking for OLPN: " + olpn);
                    } else {
                        System.err.println("❌ Failed for OLPN " + olpn + ": " + response.code());
                        System.err.println("Response: " + response.body().string());
                    }
                }
            }
        }

        private static void validateOLPNs(List<PackingData> packingList, String token, String filePath,String messageType) throws IOException {
            System.out.println("\n🔍 Validating OLPNs for MAUJDSReadyForUCSPacking = true:");

            for (PackingData data : packingList) {
                String olpnId = data.OLPNs;
                String queryJson = "{ \"Query\": \"OlpnId = " + olpnId + " AND MAUJDSReadyForUCSPacking = true\" }";

                ExcelReader reader = new ExcelReader();
                String BASE_URL             = reader.getCellValueByHeader(1, "BASE_URL");
                String SelectedOrganization = reader.getCellValueByHeader(1, "SelectedOrganization");
                String SelectedLocation     = reader.getCellValueByHeader(1, "SelectedLocation");
                reader.close();



                RequestBody body = RequestBody.create(queryJson, MediaType.get("application/json"));
                Request request = new Request.Builder()
                        .url(BASE_URL+"/pickpack/api/pickpack/olpn/search")
                        .post(body)
                        .addHeader("Authorization", "Bearer " + token)
                        .addHeader("Content-Type", "application/json")
                        .addHeader("SelectedOrganization", SelectedOrganization)
                        .addHeader("SelectedLocation", SelectedLocation)
                        .build();

                try (Response response = client.newCall(request).execute()) {
                    if (!response.isSuccessful()) {
                        System.err.println("❌ Validation failed for OLPN " + olpnId + ": " + response.code());
                        continue;
                    }

                    String responseBody = response.body().string();
                    JsonElement json = JsonParser.parseString(responseBody);
                    boolean isReady = json.toString().contains("\"MAUJDSReadyForUCSPacking\":true");

                    if (isReady) {
                        System.out.println("✅ OLPN " + olpnId + " is ready for UCS packing.");
                    } else {
                        System.err.println("⛔ OLPN " + olpnId + " is NOT ready for UCS packing. Stopping process.");
                        System.exit(1); // 🚨 Terminates the program immediately
                    }
                    System.out.println("\n⏳ Waiting 20 seconds before launching MHEJournalValidator...");
                    Thread.sleep(20000);
                    System.out.println("🚀 Launching MHEJournalValidator...\n");
                    MainA0_MHEValidator.main(filePath,messageType);
                }
                catch (InterruptedException e) {
                    throw new RuntimeException(e);
                }
            }
        }

    static String getCellValueAsString(Cell cell) {
        if (cell == null) return "";
        try {
            switch (cell.getCellType()) {
                case STRING:
                    return cell.getStringCellValue().trim();
                case NUMERIC:
                    double num = cell.getNumericCellValue();
                    return (num == (int) num) ? String.valueOf((int) num).trim() : String.valueOf(num).trim();
                case BOOLEAN:
                    return String.valueOf(cell.getBooleanCellValue()).trim();
                case FORMULA:
                    switch (cell.getCachedFormulaResultType()) {
                        case STRING:
                            return cell.getStringCellValue().trim();
                        case NUMERIC:
                            double formulaNum = cell.getNumericCellValue();
                            return (formulaNum == (int) formulaNum) ? String.valueOf((int) formulaNum).trim() : String.valueOf(formulaNum).trim();
                        case BOOLEAN:
                            return String.valueOf(cell.getBooleanCellValue()).trim();
                        default:
                            return "";
                    }
                case BLANK:
                case ERROR:
                default:
                    return "";
            }
        } catch (Exception e) {
            System.err.println("⚠️ Error reading cell value: " + e.getMessage());
            return "";
        }
    }


    static double parseDoubleSafe(Cell cell) {
        if (cell == null) return 0.0;
        try {
            switch (cell.getCellType()) {
                case NUMERIC:
                    return cell.getNumericCellValue();
                case STRING:
                    String val = cell.getStringCellValue().trim();
                    if (val.isEmpty()) return 0.0;
                    return Double.parseDouble(val);
                case FORMULA:
                    switch (cell.getCachedFormulaResultType()) {
                        case NUMERIC:
                            return cell.getNumericCellValue();
                        case STRING:
                            String fVal = cell.getStringCellValue().trim();
                            if (fVal.isEmpty()) return 0.0;
                            return Double.parseDouble(fVal);
                        default:
                            return 0.0;
                    }
                default:
                    return 0.0;
            }
        } catch (Exception e) {
            System.err.println("⚠️ Error parsing quantity: " + e.getMessage());
            return 0.0;
        }
    }










//        private static int parseIntSafe(Cell cell) throws InterruptedException {
//            try {
//                if (cell == null) return 0;
//                if (cell.getCellType() == CellType.NUMERIC) {
//                    return (int) cell.getNumericCellValue();
//                } else if (cell.getCellType() == CellType.STRING) {
//                    return Integer.parseInt(cell.getStringCellValue().trim());
//                } else if (cell.getCellType() == CellType.FORMULA) {
//                    if (cell.getCachedFormulaResultType() == CellType.NUMERIC) {
//                        return (int) cell.getNumericCellValue();
//                    } else if (cell.getCachedFormulaResultType() == CellType.STRING) {
//                        return Integer.parseInt(cell.getStringCellValue().trim());
//                    }
//                }
//            } catch (Exception e) {
//                System.err.println("⚠️ Failed to parse quantity from cell: " + e.getMessage());
//            }
//
//
//
//// ✅ Delay and launch next class
//
//
//            return 0;
//        }












    }










