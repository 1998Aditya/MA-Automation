package MA_MSG_Suite_OB;


import com.google.gson.JsonElement;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;
import okhttp3.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;


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
import java.util.ArrayList;
import java.util.List;
import java.util.UUID;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class MainA2_oLPNPrepared {

        private static final OkHttpClient client = new OkHttpClient();

        public static class OLPNSheetData {
            String WCSOrderId;
            String LCID;

            @Override
            public String toString() {
                return "WCSOrderId: " + WCSOrderId + ", LCID: " + LCID;
            }
        }

        public static void main(String filePath,String messageType) {
            try {
                List<OLPNSheetData> olpnList = readOlpnData(filePath);
                System.out.println("✅ oLPNPrepared Data Extracted:");
                olpnList.forEach(System.out::println);

                System.out.println("\n📦 Generated JSON Payloads:");
                for (OLPNSheetData data : olpnList) {
                    JsonObject payload = buildPayload(data,messageType);
                    System.out.println("------------------------------------------------------");
                    System.out.println(payload.toString());
                }

                String token = getAuthTokenFromExcel();
                if (token == null) {
                    System.err.println("❌ Failed to authenticate.");
                    return;
                }

                triggerAPI(olpnList, token,messageType);           // ✅ First, post the message
                validatePackedStatus(olpnList, token,filePath,messageType); // ⛔ Then validate packing status

            } catch (Exception e) {
                System.err.println("❌ Error: " + e.getMessage());
                e.printStackTrace();
            }
        }

        private static List<OLPNSheetData> readOlpnData(String path) throws IOException {
            List<OLPNSheetData> list = new ArrayList<>();

            try (FileInputStream fis = new FileInputStream(path);
                 Workbook workbook = new XSSFWorkbook(fis)) {

                Sheet sheet = workbook.getSheet("Tasks");
                if (sheet == null) {
                    System.err.println("❌ Sheet 'Tasks' not found.");
                    return list;
                }

                for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                    Row row = sheet.getRow(i);
                    if (row == null) continue;

                    OLPNSheetData data = new OLPNSheetData();
                    data.WCSOrderId = getCellValueAsString(row.getCell(4));
                    data.LCID = getCellValueAsString(row.getCell(19));

                    if (!data.WCSOrderId.isEmpty() && !data.LCID.isEmpty()) {
                        list.add(data);
                    }
                }
            }
            return list;
        }

        private static JsonObject buildPayload(OLPNSheetData data,String messageType) {
            JsonObject json = new JsonObject();
            json.addProperty("WCSOrderId", data.WCSOrderId);
            json.addProperty("LCID", data.LCID);
            json.addProperty("MessageType", messageType);
            json.addProperty("UniqueKey", UUID.randomUUID().toString().replace("-", ""));
            return json;
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

        private static void triggerAPI(List<OLPNSheetData> olpnList, String token,String messageType) throws IOException {
            for (OLPNSheetData data : olpnList) {
                JsonObject payload = buildPayload(data,messageType);

                ExcelReader reader = new ExcelReader();
                String BASE_URL             = reader.getCellValueByHeader(1, "BASE_URL");
                String SelectedOrganization = reader.getCellValueByHeader(1, "SelectedOrganization");
                String SelectedLocation     = reader.getCellValueByHeader(1, "SelectedLocation");
                reader.close();


                RequestBody body = RequestBody.create(payload.toString(), MediaType.get("application/json"));
                Request request = new Request.Builder()
                        .url(BASE_URL+"device-integration/api/deviceintegration/process/oLPNPrepared_FER_Src_EP")
                        .post(body)
                        .addHeader("Authorization", "Bearer " + token)
                        .addHeader("Content-Type", "application/json")
                        .addHeader("SelectedOrganization", SelectedOrganization)
                        .addHeader("SelectedLocation", SelectedLocation)
                        .build();

                try (Response response = client.newCall(request).execute()) {
                    if (response.isSuccessful()) {
                        System.out.println("✅ Sent oLPNPrepared for WCSOrderId: " + data.WCSOrderId);
                    } else {
                        System.err.println("❌ Failed for WCSOrderId " + data.WCSOrderId + ": " + response.code());
                        System.err.println("Response: " + response.body().string());
                    }
                }
            }
        }

        private static void validatePackedStatus(List<OLPNSheetData> olpnList, String token,String filePath,String messageType) throws IOException {
            System.out.println("\n🔍 Validating OLPNs for Status = 7200 (Packed):");

            for (OLPNSheetData data : olpnList) {
                String olpnId = data.WCSOrderId;
                String queryJson = "{ \"Query\": \"OlpnId = '" + olpnId + "' AND Status = 7200\" }";

                ExcelReader reader = new ExcelReader();
                String BASE_URL             = reader.getCellValueByHeader(1, "BASE_URL");
                String SelectedOrganization = reader.getCellValueByHeader(1, "SelectedOrganization");
                String SelectedLocation     = reader.getCellValueByHeader(1, "SelectedLocation");
                reader.close();



                RequestBody body = RequestBody.create(queryJson, MediaType.get("application/json"));
                Request request = new Request.Builder()
                        .url(BASE_URL+"pickpack/api/pickpack/olpn/search")
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
                    System.out.println("🔍 Raw Response for OLPN " + olpnId + ": " + responseBody);

                    JsonElement json = JsonParser.parseString(responseBody);
                    boolean found = false;

                    if (json.isJsonObject()) {
                        JsonObject root = json.getAsJsonObject();
                        JsonElement dataElement = root.has("data") ? root.get("data") : null;

                        if (dataElement != null && dataElement.isJsonArray()) {
                            found = dataElement.getAsJsonArray().size() > 0;
                        }
                    }

                    if (found) {
                        System.out.println("✅ OLPN " + olpnId + " is packed.");
                    } else {
                        System.err.println("⛔ OLPN " + olpnId + " is NOT packed. Stopping process.");
                        System.exit(1);
                    }
                    System.out.println("\n⏳ Waiting 20 seconds before launching MHEJournalValidator...");
                    Thread.sleep(20000);
                    System.out.println("🚀 Launching MHEJournalValidator...\n");

                    MainA0_MHEValidator.main(filePath,messageType);


                } catch (InterruptedException e) {
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

        private static int parseIntSafe(Cell cell) {
            try {
                if (cell == null) return 0;
                if (cell.getCellType() == CellType.NUMERIC) {
                    return (int) cell.getNumericCellValue();
                } else if (cell.getCellType() == CellType.STRING) {
                    return Integer.parseInt(cell.getStringCellValue().trim());
                } else if (cell.getCellType() == CellType.FORMULA) {
                    if (cell.getCachedFormulaResultType() == CellType.NUMERIC) {
                        return (int) cell.getNumericCellValue();
                    } else if (cell.getCachedFormulaResultType() == CellType.STRING) {
                        return Integer.parseInt(cell.getStringCellValue().trim());
                    }
                }
            } catch (Exception e) {
                System.err.println("⚠️ Failed to parse quantity from cell: " + e.getMessage());
            }
            return 0;

        }
    }

