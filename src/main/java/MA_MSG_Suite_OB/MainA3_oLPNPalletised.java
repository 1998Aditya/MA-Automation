package MA_MSG_Suite_OB;



import com.google.gson.JsonObject;
import okhttp3.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;
import com.google.gson.JsonParser;
import com.google.gson.JsonElement;
import com.google.gson.JsonElement;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;
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
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class MainA3_oLPNPalletised {


        private static final OkHttpClient client = new OkHttpClient();


        public static class PalletisedSheetData {
            String WCSOrderId;
            String PalletId;
            String LocationId;

            @Override
            public String toString() {
                return "WCSOrderId: " + WCSOrderId + ", PalletId: " + PalletId + ", LocationId: " + LocationId;
            }
        }

        public static void main(String filePath, String messageType) {
            try {
                List<PalletisedSheetData> palletList = readPalletisedData(filePath);
                System.out.println("✅ Extracted oLPNPalletised Data:");
                palletList.forEach(System.out::println);

                String token = getAuthTokenFromExcel();
                if (token == null) {
                    System.err.println("❌ Authentication failed.");
                    return;
                }

                triggerAPI(palletList, token,filePath,messageType);

            } catch (Exception e) {
                System.err.println("❌ Error: " + e.getMessage());
                e.printStackTrace();
            }
        }

        private static List<PalletisedSheetData> readPalletisedData(String path) throws IOException {
            List<PalletisedSheetData> list = new ArrayList<>();

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

                    PalletisedSheetData data = new PalletisedSheetData();
                    data.WCSOrderId = getCellValueAsString(row.getCell(9));   // OLPN
                    data.PalletId = getCellValueAsString(row.getCell(20));    // PalletId
                    data.LocationId = getCellValueAsString(row.getCell(21));  // LocationId

                    if (!data.WCSOrderId.isEmpty() && !data.PalletId.isEmpty() && !data.LocationId.isEmpty()) {
                        list.add(data);
                    }
                }
            }
            return list;
        }

        private static JsonObject buildPayload(PalletisedSheetData data) {
            JsonObject json = new JsonObject();
            json.addProperty("WCSOrderId", data.WCSOrderId);
            json.addProperty("PalletId", data.PalletId);
            json.addProperty("LocationId", data.LocationId);
            json.addProperty("MessageType", "oLPNPalletised");
            json.addProperty("UniqueKey", String.valueOf(System.currentTimeMillis()));
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

    private static void triggerAPI(List<PalletisedSheetData> palletList, String token,String filePath,String messageType) throws IOException {
            for (PalletisedSheetData data : palletList) {
                JsonObject payload = buildPayload(data);

                System.out.println("\n📤 Sending Payload for WCSOrderId: " + data.WCSOrderId);
                System.out.println(payload.toString());


                ExcelReader reader = new ExcelReader();
                String BASE_URL             = reader.getCellValueByHeader(1, "BASE_URL");
                String SelectedOrganization = reader.getCellValueByHeader(1, "SelectedOrganization");
                String SelectedLocation     = reader.getCellValueByHeader(1, "SelectedLocation");
                reader.close();



                RequestBody body = RequestBody.create(payload.toString(), MediaType.get("application/json"));
                Request request = new Request.Builder()
                        .url(BASE_URL+"/device-integration/api/deviceintegration/process/oLPNPalletised_FER_Src_EP")
                        .post(body)
                        .addHeader("Authorization", "Bearer " + token)
                        .addHeader("Content-Type", "application/json")
                        .addHeader("SelectedOrganization", SelectedOrganization)
                        .addHeader("SelectedLocation", SelectedLocation)
                        .build();

                try (Response response = client.newCall(request).execute()) {
                    String responseBody = response.body().string();
                    System.out.println("🔍 Response Code: " + response.code());
                    System.out.println("🔍 Response Body: " + responseBody);

                    if (response.isSuccessful()) {
                        System.out.println("✅ Successfully posted oLPNPalletised for WCSOrderId: " + data.WCSOrderId);

                        // 🔍 Validate PalletId
                        boolean isValid = validatePalletId(data.WCSOrderId, data.PalletId, token,filePath,messageType);
                        if (!isValid) {
                            System.err.println("🛑 Halting process due to PalletId mismatch.");
                            return;
                        }

                    } else {
                        System.err.println("❌ Failed for WCSOrderId " + data.WCSOrderId);
                        return;
                    }
                }
            }
        }

        private static boolean validatePalletId(String olpnId, String expectedPalletId, String token,String filePath,String messageType) throws IOException {
            // ✅ Build correct query with quoted OLPN and PalletId condition
            JsonObject query = new JsonObject();
            query.addProperty("Query", "OlpnId = '" + olpnId + "' AND PalletId != ''");

            ExcelReader reader = new ExcelReader();
            String BASE_URL             = reader.getCellValueByHeader(1, "BASE_URL");
            String SelectedOrganization = reader.getCellValueByHeader(1, "SelectedOrganization");
            String SelectedLocation     = reader.getCellValueByHeader(1, "SelectedLocation");
            reader.close();



            RequestBody body = RequestBody.create(query.toString(), MediaType.get("application/json"));
            Request request = new Request.Builder()
                    .url(BASE_URL)
                    .post(body)
                    .addHeader("Authorization", "Bearer " + token)
                    .addHeader("Content-Type", "application/json")
                    .addHeader("SelectedOrganization", SelectedOrganization)
                    .addHeader("SelectedLocation", SelectedLocation)
                    .build();

            try (Response response = client.newCall(request).execute()) {
                String responseBody = response.body().string();
                System.out.println("🔍 Validation Response: " + responseBody);

                if (!response.isSuccessful()) {
                    System.err.println("❌ Validation API failed for OLPN: " + olpnId);
                    return false;
                }

                JsonObject json = JsonParser.parseString(responseBody).getAsJsonObject();
                JsonElement dataElement = json.has("data") ? json.get("data") : null;

                if (dataElement == null || !dataElement.isJsonArray() || dataElement.getAsJsonArray().size() == 0) {
                    System.err.println("❌ No 'data' found for OLPN: " + olpnId);
                    return false;
                }

                JsonObject result = dataElement.getAsJsonArray().get(0).getAsJsonObject();
                String actualPalletId = result.has("PalletId") ? result.get("PalletId").getAsString() : "";

                boolean match = expectedPalletId.equals(actualPalletId);
                if (!match) {
                    System.err.println("❌ PalletId mismatch for OLPN " + olpnId + ". Expected: " + expectedPalletId + ", Found: " + actualPalletId);
                } else {
                    System.out.println("✅ OLPN " + olpnId + " has correct PalletId.");
                }
                System.out.println("\n⏳ Waiting 20 seconds before launching MHEJournalValidator...");
                Thread.sleep(20000);
                System.out.println("🚀 Launching MHEJournalValidator...\n");
                MainA0_MHEValidator.main(filePath,messageType);
                return match;
            } catch (InterruptedException e) {
                throw new RuntimeException(e);
            }
        }


        private static String getCellValueAsString(Cell cell) {
            if (cell == null) return "";
            switch (cell.getCellType()) {
                case STRING: return cell.getStringCellValue().trim();
                case NUMERIC: return String.valueOf((long) cell.getNumericCellValue()).trim();
                case BOOLEAN: return String.valueOf(cell.getBooleanCellValue()).trim();
                case FORMULA: return cell.getCellFormula().trim();
                default: return "";
            }
        }
    }
