//package MA_MSG_Suite_OB;
//
//
//import com.google.gson.JsonObject;
//import com.google.gson.JsonParser;
//import okhttp3.*;
//import org.apache.poi.ss.usermodel.*;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//
//import java.io.FileInputStream;
//import java.io.IOException;
//import java.util.*;
//import com.google.gson.JsonObject;
//import io.github.bonigarcia.wdm.WebDriverManager;
//import okhttp3.*;
//import org.apache.poi.ss.usermodel.Cell;
//import org.apache.poi.ss.usermodel.Row;
//import org.apache.poi.ss.usermodel.Sheet;
//import org.apache.poi.ss.usermodel.Workbook;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//import org.openqa.selenium.*;
//import org.openqa.selenium.chrome.ChromeDriver;
//import org.openqa.selenium.chrome.ChromeOptions;
//import org.openqa.selenium.support.ui.ExpectedConditions;
//import org.openqa.selenium.support.ui.WebDriverWait;
//
//import java.io.FileInputStream;
//import java.io.IOException;
//import java.time.Duration;
//import java.util.ArrayList;
//import java.util.List;
//import java.util.regex.Matcher;
//import java.util.regex.Pattern;
//
//public class MainA4_PalletReady {
//
//
//
//        private static final OkHttpClient client = new OkHttpClient();
//
//
//        public static class PalletReadyData {
//            String PalletId;
//
//            @Override
//            public String toString() {
//                return "PalletId: " + PalletId;
//            }
//        }
//
//        public static void main(String filePath,String messageType) {
//            try {
//                List<PalletReadyData> palletList = readPalletReadyData(filePath);
//                System.out.println("✅ Extracted PalletReady Data:");
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
//        private static List<PalletReadyData> readPalletReadyData(String path) throws IOException {
//            List<PalletReadyData> list = new ArrayList<>();
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
//                    PalletReadyData data = new PalletReadyData();
//                    data.PalletId = getCellValueAsString(row.getCell(20));  // PalletId from column 20
//
//                    if (!data.PalletId.isEmpty()) {
//                        list.add(data);
//                    }
//                }
//            }
//            return list;
//        }
//
//        private static JsonObject buildPayload(PalletReadyData data) {
//            JsonObject json = new JsonObject();
//            json.addProperty("PalletId", data.PalletId);
//            json.addProperty("MessageType", "PalletReady");
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
//
//
//    private static void triggerAPI(List<PalletReadyData> palletList, String token,String filePath,String messageType) throws IOException {
//            for (PalletReadyData data : palletList) {
//                JsonObject payload = buildPayload(data);
//
//                System.out.println("\n📤 Sending Payload for PalletId: " + data.PalletId);
//                System.out.println(payload.toString());
//
//                ExcelReader reader = new ExcelReader();
//                String BASE_URL             = reader.getCellValueByHeader(1, "BASE_URL");
//                String SelectedOrganization = reader.getCellValueByHeader(1, "SelectedOrganization");
//                String SelectedLocation     = reader.getCellValueByHeader(1, "SelectedLocation");
//                reader.close();
//
//
//
//
//
//                RequestBody body = RequestBody.create(payload.toString(), MediaType.get("application/json"));
//                Request request = new Request.Builder()
//                        .url(BASE_URL+"/device-integration/api/deviceintegration/process/PalletReady_FER_Src_EP")
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
//                        System.out.println("✅ Successfully posted PalletReady for PalletId: " + data.PalletId);
//                    } else {
//                        System.err.println("❌ Failed for PalletId " + data.PalletId);
//                    }
//                    System.out.println("\n⏳ Waiting 20 seconds before launching MHEJournalValidator...");
//                    Thread.sleep(20000);
//                    System.out.println("🚀 Launching MHEJournalValidator...\n");
//                    MainA0_MHEValidator.main(filePath,messageType);
//                } catch (InterruptedException e) {
//                    throw new RuntimeException(e);
//                }
//            }
//        }
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
//





package MA_MSG_Suite_OB;

import com.google.gson.JsonObject;
import com.google.gson.JsonParser;
import okhttp3.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;
import org.openqa.selenium.*;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.time.Duration;

public class MainA4_PalletReady {
    public static String docPathLocal ;

    private static final OkHttpClient client = new OkHttpClient();

    public static class PalletReadyData {
        String PalletId;

        @Override
        public String toString() {
            return "PalletId: " + PalletId;
        }
    }

    // NOTE: Keep the signature — you already have testcase and messageType available
    public static void main(String filePath, String testcase, String messageType,String env) {
        docPathLocal = DocPathManager.getOrCreateDocPath(filePath, testcase);

        try {
            // ✅ Now filters by `testcase` exactly like your previous implementation
            List<PalletReadyData> palletList = readPalletReadyData(filePath, testcase);

            System.out.println("✅ Extracted PalletReady Data (filtered by Testcase): " + testcase);
            palletList.forEach(System.out::println);

            if (palletList.isEmpty()) {
                System.err.println("⚠️ No rows found for Testcase: " + testcase + " in sheet 'Tasks'.");
                return;
            }

            String token = getAuthTokenFromExcel();
            if (token == null) {
                System.err.println("❌ Authentication failed.");
                return;
            }

            triggerAPI(palletList, token, filePath, messageType);
            Thread.sleep(5000);
            Main100Pallet_MHEJournalScreenshot.main(testcase,filePath, env, messageType,docPathLocal);
            Thread.sleep(5000);

        } catch (Exception e) {
            System.err.println("❌ Error: " + e.getMessage());
            e.printStackTrace();
        }
    }



//    private static List<PalletReadyData> readPalletReadyData(String path, String testcase) throws IOException {
//        List<PalletReadyData> list = new ArrayList<>();
//
//        try (FileInputStream fis = new FileInputStream(path);
//             Workbook workbook = new XSSFWorkbook(fis)) {
//
//            Sheet sheet = workbook.getSheet("Tasks");
//            if (sheet == null) {
//                System.err.println("❌ Sheet 'Tasks' not found.");
//                return list;
//            }
//
//            DataFormatter fmt = new DataFormatter();
//
//            // ✅ Find the column index for "Pallet"
//            int palletCol = getColumnIndexByName(sheet, "Pallet");
//            if (palletCol == -1) {
//                throw new RuntimeException("❌ Column 'Pallet' not found in Tasks sheet header.");
//            }
//
//            // Testcase filter logic
//            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
//                Row row = sheet.getRow(i);
//                if (row == null) continue;
//
//                String testcaseCell = fmt.formatCellValue(row.getCell(0)).trim();
//                if (!testcaseCell.equalsIgnoreCase(testcase.trim())) continue;
//
//                // ✅ Read using column name
//                String palletId = fmt.formatCellValue(row.getCell(palletCol)).trim();
//
//                if (!palletId.isEmpty()) {
//                    PalletReadyData data = new PalletReadyData();
//                    data.PalletId = palletId;
//                    list.add(data);
//                }
//            }
//        }
//        return list;
//    }




    private static List<PalletReadyData> readPalletReadyData(String path, String testcase) throws IOException {
        List<PalletReadyData> list = new ArrayList<>();
        Set<String> uniquePallets = new HashSet<>(); // ✅ Track uniqueness

        try (FileInputStream fis = new FileInputStream(path);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheet("Tasks");
            if (sheet == null) {
                System.err.println("❌ Sheet 'Tasks' not found.");
                return list;
            }

            DataFormatter fmt = new DataFormatter();

            // ✅ Find the column index for "Pallet"
            int palletCol = getColumnIndexByName(sheet, "Pallet");
            if (palletCol == -1) {
                throw new RuntimeException("❌ Column 'Pallet' not found in Tasks sheet header.");
            }

            // Testcase filter logic
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                String testcaseCell = fmt.formatCellValue(row.getCell(0)).trim();
                if (!testcaseCell.equalsIgnoreCase(testcase.trim())) continue;

                // Read using column name
                String palletId = fmt.formatCellValue(row.getCell(palletCol)).trim();

                if (!palletId.isEmpty() && uniquePallets.add(palletId)) {
                    // add() returns false if palletId already exists
                    PalletReadyData data = new PalletReadyData();
                    data.PalletId = palletId;
                    list.add(data);
                }
            }
        }
        return list;
    }





    private static int getColumnIndexByName(Sheet sheet, String columnName) {
        Row headerRow = sheet.getRow(0);  // header is in row 0
        if (headerRow == null) return -1;

        DataFormatter formatter = new DataFormatter();

        for (int i = 0; i < headerRow.getLastCellNum(); i++) {
            Cell cell = headerRow.getCell(i);
            if (cell == null) continue;

            String header = formatter.formatCellValue(cell).trim();
            if (header.equalsIgnoreCase(columnName.trim())) {
                return i;   // found!
            }
        }
        return -1; // not found
    }

    // ✅ Accepts messageType from caller instead of hardcoding "PalletReady"
    private static JsonObject buildPayload(PalletReadyData data, String messageType) {
        JsonObject json = new JsonObject();
        json.addProperty("PalletId", data.PalletId);
        json.addProperty("MessageType", (messageType == null || messageType.isEmpty()) ? "PalletReady" : messageType);
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

    private static void triggerAPI(List<PalletReadyData> palletList, String token, String filePath, String messageType) throws IOException {
        for (PalletReadyData data : palletList) {
            JsonObject payload = buildPayload(data, messageType);

            System.out.println("\n📤 Sending Payload for PalletId: " + data.PalletId);
            System.out.println(payload.toString());

            ExcelReader reader = new ExcelReader();
            String BASE_URL             = reader.getCellValueByHeader(1, "BASE_URL");
            String SelectedOrganization = reader.getCellValueByHeader(1, "SelectedOrganization");
            String SelectedLocation     = reader.getCellValueByHeader(1, "SelectedLocation");
            reader.close();

            RequestBody body = RequestBody.create(payload.toString(), MediaType.get("application/json"));
            Request request = new Request.Builder()
                    .url(BASE_URL + "/device-integration/api/deviceintegration/process/PalletReady_FER_Src_EP")
                    .post(body)
                    .addHeader("Authorization", "Bearer " + token)
                    .addHeader("Content-Type", "application/json")
                    .addHeader("SelectedOrganization", SelectedOrganization)
                    .addHeader("SelectedLocation", SelectedLocation)
                    .build();

            try (Response response = client.newCall(request).execute()) {
                String responseBody = response.body() != null ? response.body().string() : "";
                System.out.println("🔍 Response Code: " + response.code());
                System.out.println("🔍 Response Body: " + responseBody);

                if (response.isSuccessful()) {
                    System.out.println("✅ Successfully posted PalletReady for PalletId: " + data.PalletId);
                } else {
                    System.err.println("❌ Failed for PalletId " + data.PalletId);
                }

//                System.out.println("\n⏳ Waiting 20 seconds before launching MHEJournalValidator...");
//                try {
//                    Thread.sleep(20000);
//                } catch (InterruptedException ie) {
//                    Thread.currentThread().interrupt();
//                    throw new RuntimeException(ie);
//                }
//                System.out.println("🚀 Launching MHEJournalValidator...\n");
             //   MainA0_MHEValidator.main(filePath, messageType);
            }
        }
    }
}
