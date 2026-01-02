package MA_MSG_Suite_OB;

import com.google.gson.*;
import com.google.gson.stream.JsonReader;
import okhttp3.*;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class Main6b_WaveToExcel {

   // public static String waveNumber = "W20251116231500000001";
    public static WebDriver driver;
  //  public static String EXCEL_PATH = "C:\\Users\\2210420\\IdeaProjects\\Testcases\\OOdata.xlsx";
//    public static Map<String, List<String>> waveTaskMap = new HashMap<>();
//    public static Map<String, List<String>> waveOlpnMap = new HashMap<>();

    public static void main(String waveNumber,String EXCEL_PATH, String testcase) throws InterruptedException, IOException {
        try {
            processWaveRunData(waveNumber,EXCEL_PATH,testcase);
        } catch (Exception e) {
            System.err.println("❌ Error occurred in main: " + e.getMessage());
            e.printStackTrace();
        }
    }


    private static void processWaveRunData(String waveRun,String EXCEL_PATH,String testcase) throws Exception {
        String token = getAuthTokenFromExcel();
        if (token == null) {
            System.err.println("❌ Failed to retrieve access token.");
            return;
        }

        // ✅ Close Excel if open
        closeExcelIfOpen();

        try (FileInputStream fis = new FileInputStream(EXCEL_PATH);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(3);

            // ✅ Create header if not present
            if (sheet.getLastRowNum() == 0 && sheet.getRow(0) == null) {
                Row header = sheet.createRow(0);
                header.createCell(0).setCellValue("testcases");
                header.createCell(1).setCellValue("WaveRun");
                header.createCell(2).setCellValue("TaskId");
                header.createCell(3).setCellValue("TaskGroup");
                header.createCell(4).setCellValue("OLPNId");
                header.createCell(5).setCellValue("Location");
                header.createCell(6).setCellValue("Item");
                header.createCell(7).setCellValue("Quantity");
            }
//
//            // ✅ Clear old data except header
//            int lastRow = sheet.getLastRowNum();
//            for (int i = lastRow; i >= 1; i--) {
//                Row row = sheet.getRow(i);
//                if (row != null) sheet.removeRow(row);
//            }

            // ✅ Fetch OLPN IDs for WaveRun
            List<String> olpnIds = fetchIds(waveRun, "olpn", "OrderPlanningRunId", token, "com-manh-cp-pickpack", "OlpnId");
            System.out.println("✅ WaveRun: " + waveRun + " | OLPN IDs: " + olpnIds.size());

            // ✅ Start writing after the last existing row
            int rowIndex = sheet.getLastRowNum() + 1;

            for (String olpnId : olpnIds) {
                // ✅ Fetch TaskId for OLPN using new API
                List<String> taskIds = fetchTaskIdsByOlpn(olpnId, token);
                String taskId = taskIds.isEmpty() ? "" : taskIds.get(0);

                // ✅ Fetch TaskGroup for TaskId
                String taskGroup = taskId.isEmpty() ? "" : fetchTaskGroup(taskId, token);

                // ✅ Write OLPN + TaskId + TaskGroup
                Row row = sheet.createRow(rowIndex++);
                row.createCell(0).setCellValue(testcase);
                row.createCell(1).setCellValue(waveRun);
                row.createCell(2).setCellValue(taskId);
                row.createCell(3).setCellValue(taskGroup);
                row.createCell(4).setCellValue(olpnId);

                // ✅ Fetch Task Details if TaskId exists
                if (!taskId.isEmpty()) {
                    List<WavingToExcel.TaskDetail> details = fetchTaskDetails(taskId, token);
                    writeTaskDetails(row, details);
                }
            }

            try (FileOutputStream fos = new FileOutputStream(EXCEL_PATH)) {
                workbook.write(fos);
            }

            System.out.println("✅ WaveRun, OLPN, TaskId, TaskGroup, and Task Details written successfully.");
            System.out.println("-------------------------------------------MainC2_Wave-To-Excel- complete --------------------------");

        }
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
    // ✅ Fetch Task IDs by OLPN
    public static List<String> fetchTaskIdsByOlpn(String olpnId, String token) throws IOException {
        ExcelReader reader = new ExcelReader();

// By header name
        String BASE_URL = reader.getCellValueByHeader(1, "BASE_URL");
        String SelectedOrganization = reader.getCellValueByHeader(1, "SelectedOrganization");
        String SelectedLocation = reader.getCellValueByHeader(1, "SelectedLocation");

        reader.close();


        List<String> taskIds = new ArrayList<>();
        try {
            OkHttpClient client = new OkHttpClient();
            MediaType mediaType = MediaType.parse("application/json");

            JsonObject filter = new JsonObject();
            filter.addProperty("ViewName", "task");
            filter.addProperty("AttributeId", "TaskDetail.OlpnId");
            filter.addProperty("Operator", "=");
            filter.add("FilterValues", new Gson().toJsonTree(List.of(olpnId)));

            JsonObject body = new JsonObject();
            body.addProperty("ViewName", "Task");
            body.add("Filters", new Gson().toJsonTree(List.of(filter)));
            body.addProperty("ComponentName", "com-manh-cp-task");
            body.addProperty("Size", 25);
            body.addProperty("Sort", "CreatedDateTime");
            body.addProperty("SortOrder", "desc");
            body.addProperty("TimeZone", "Europe/Paris");

            RequestBody requestBody = RequestBody.create(mediaType, body.toString());
            Request request = new Request.Builder()
                    .url(BASE_URL+"/dmui-facade/api/dmui-facade/entity/search")
                    .post(requestBody)
                    .addHeader("Content-Type", "application/json")
                    .addHeader("Authorization", "Bearer " + token)
                    .addHeader("SelectedOrganization", SelectedOrganization)
                    .addHeader("SelectedLocation", SelectedLocation)
                    .build();

            Response response = client.newCall(request).execute();
            String responseBody = response.body() != null ? response.body().string() : "";

            if (!responseBody.trim().startsWith("{")) return taskIds;

            JsonObject json = JsonParser.parseString(responseBody).getAsJsonObject();
            if (json.has("data")) {
                JsonArray results = json.getAsJsonObject("data").getAsJsonArray("Results");
                for (JsonElement element : results) {
                    JsonObject obj = element.getAsJsonObject();
                    if (obj.has("TaskId")) taskIds.add(obj.get("TaskId").getAsString());
                }
            }
        } catch (Exception e) {
            System.err.println("❌ Error fetching Task IDs by OLPN: " + e.getMessage());
        }
        return taskIds;
    }
    // ✅ Fetch TaskGroup
    private static String fetchTaskGroup(String taskId, String token) throws IOException {
        ExcelReader reader = new ExcelReader();

// By header name
        String BASE_URL = reader.getCellValueByHeader(1, "BASE_URL");
        String SelectedOrganization = reader.getCellValueByHeader(1, "SelectedOrganization");
        String SelectedLocation = reader.getCellValueByHeader(1, "SelectedLocation");

        reader.close();
        String taskGroup = "";
        try {
            OkHttpClient client = new OkHttpClient();
            MediaType mediaType = MediaType.parse("application/json");

            JsonObject filter = new JsonObject();
            filter.addProperty("ViewName", "AssignableTaskGroup");
            filter.addProperty("AttributeId", "TaskId");
            filter.addProperty("Operator", "=");
            filter.add("FilterValues", new Gson().toJsonTree(List.of(taskId)));

            JsonObject body = new JsonObject();
            body.addProperty("ViewName", "AssignableTaskGroup");
            body.add("Filters", new Gson().toJsonTree(List.of(filter)));
            body.addProperty("ComponentName", "com-manh-cp-task");
            body.addProperty("Size", 25);
            body.addProperty("TimeZone", "Europe/Paris");

            RequestBody requestBody = RequestBody.create(mediaType, body.toString());
            Request request = new Request.Builder()
                    .url(BASE_URL+"/dmui-facade/api/dmui-facade/entity/search")
                    .post(requestBody)
                    .addHeader("Content-Type", "application/json")
                    .addHeader("Authorization", "Bearer " + token)
                    .addHeader("SelectedOrganization", SelectedOrganization)
                    .addHeader("SelectedLocation", SelectedLocation)
                    .build();

            Response response = client.newCall(request).execute();
            String responseBody = response.body() != null ? response.body().string() : "";

            if (!responseBody.trim().startsWith("{")) return taskGroup;

            JsonObject json = JsonParser.parseString(responseBody).getAsJsonObject();
            if (json.has("data")) {
                JsonArray results = json.getAsJsonObject("data").getAsJsonArray("Results");
                if (results.size() > 0) {
                    JsonObject obj = results.get(0).getAsJsonObject();
                    taskGroup = obj.has("TaskGroupId") ? obj.get("TaskGroupId").getAsString() : "";
                }
            }
        } catch (Exception e) {
            System.err.println("❌ Error fetching TaskGroup: " + e.getMessage());
        }
        return taskGroup;
    }
    // ✅ Write Task Details starting from column F
    private static void writeTaskDetails(Row row, List<WavingToExcel.TaskDetail> details) {
        int colIndex = 7; // Column H
        for (WavingToExcel.TaskDetail detail : details) {
            row.createCell(colIndex).setCellValue(detail.getLocation());
            row.createCell(colIndex + 1).setCellValue(detail.getItemId());
            row.createCell(colIndex + 2).setCellValue(detail.getQuantity());
            colIndex += 3;
        }
    }
    // ✅ Fetch Task Details
    private static List<WavingToExcel.TaskDetail> fetchTaskDetails(String taskId, String token) throws IOException {
        ExcelReader reader = new ExcelReader();

// By header name
        String BASE_URL = reader.getCellValueByHeader(1, "BASE_URL");
        String SelectedOrganization = reader.getCellValueByHeader(1, "SelectedOrganization");
        String SelectedLocation = reader.getCellValueByHeader(1, "SelectedLocation");

        reader.close();
        List<WavingToExcel.TaskDetail> details = new ArrayList<>();
        try {
            OkHttpClient client = new OkHttpClient();
            MediaType mediaType = MediaType.parse("application/json");

            JsonObject filter = new JsonObject();
            filter.addProperty("ViewName", "taskDetail");
            filter.addProperty("AttributeId", "TaskId");
            filter.addProperty("Operator", "=");
            filter.add("FilterValues", new Gson().toJsonTree(List.of(taskId)));

            JsonObject body = new JsonObject();
            body.addProperty("ViewName", "TaskDetail");
            body.add("Filters", new Gson().toJsonTree(List.of(filter)));
            body.addProperty("ComponentName", "com-manh-cp-task");
            body.addProperty("Size", 25);
            body.addProperty("Sort", "CreatedTimestamp");
            body.addProperty("SortOrder", "desc");
            body.addProperty("TimeZone", "Europe/Paris");

            RequestBody requestBody = RequestBody.create(mediaType, body.toString());
            Request request = new Request.Builder()
                    .url(BASE_URL+"/dmui-facade/api/dmui-facade/entity/search")
                    .post(requestBody)
                    .addHeader("Content-Type", "application/json")
                    .addHeader("Authorization", "Bearer " + token)
                    .addHeader("SelectedOrganization", SelectedOrganization)
                    .addHeader("SelectedLocation", SelectedLocation)
                    .build();

            Response response = client.newCall(request).execute();
            String responseBody = response.body() != null ? response.body().string() : "";

            if (!responseBody.trim().startsWith("{")) return details;

            JsonReader reader1 = new JsonReader(new StringReader(responseBody));
            reader1.setLenient(true);
            JsonObject json = JsonParser.parseReader(reader1).getAsJsonObject();

            if (json.has("data")) {
                JsonArray results = json.getAsJsonObject("data").getAsJsonArray("Results");
                for (JsonElement element : results) {
                    JsonObject obj = element.getAsJsonObject();
                    String itemId = obj.has("ItemId") ? obj.get("ItemId").getAsString() : "";
                    String location = obj.has("SourceLocationId") ? obj.get("SourceLocationId").getAsString() : "";
                    String quantity = obj.has("Quantity") ? obj.get("Quantity").getAsString() : "";
                    details.add(new WavingToExcel.TaskDetail(itemId, quantity, location));
                }
            }
        } catch (Exception e) {
            System.err.println("❌ Error fetching task details: " + e.getMessage());
        }
        return details;
    }
    // ✅ Fetch IDs (OLPN)
    public static List<String> fetchIds(String waveRun, String viewName, String attributeId, String token, String componentName, String idKey) throws IOException {
        ExcelReader reader = new ExcelReader();

// By header name
        String BASE_URL = reader.getCellValueByHeader(1, "BASE_URL");
        String SelectedOrganization = reader.getCellValueByHeader(1, "SelectedOrganization");
        String SelectedLocation = reader.getCellValueByHeader(1, "SelectedLocation");

        reader.close();

        List<String> ids = new ArrayList<>();
        try {
            OkHttpClient client = new OkHttpClient();
            MediaType mediaType = MediaType.parse("application/json");

            JsonObject filter = new JsonObject();
            filter.addProperty("ViewName", viewName);
            filter.addProperty("AttributeId", attributeId);
            filter.addProperty("Operator", "=");
            filter.add("FilterValues", new Gson().toJsonTree(List.of(waveRun.trim())));

            JsonObject body = new JsonObject();
            body.addProperty("ViewName", viewName.equalsIgnoreCase("olpn") ? "DMOlpn" : "Task");
            body.add("Filters", new Gson().toJsonTree(List.of(filter)));
            body.addProperty("ComponentName", componentName);
            body.addProperty("Size", 100);
            body.addProperty("TimeZone", "Europe/Paris");

            RequestBody requestBody = RequestBody.create(mediaType, body.toString());
            Request request = new Request.Builder()
                    .url(BASE_URL+"/dmui-facade/api/dmui-facade/entity/search")
                    .post(requestBody)
                    .addHeader("Content-Type", "application/json")
                    .addHeader("Authorization", "Bearer " + token)
                    .addHeader("SelectedOrganization", SelectedOrganization)
                    .addHeader("SelectedLocation", SelectedLocation)
                    .build();

            Response response = client.newCall(request).execute();
            String responseBody = response.body() != null ? response.body().string() : "";

            if (!responseBody.trim().startsWith("{")) return ids;

            JsonReader reader1 = new JsonReader(new StringReader(responseBody));
            reader1.setLenient(true);
            JsonObject json = JsonParser.parseReader(reader1).getAsJsonObject();

            if (json.has("data")) {
                JsonArray results = json.getAsJsonObject("data").getAsJsonArray("Results");
                for (JsonElement element : results) {
                    JsonObject obj = element.getAsJsonObject();
                    if (obj.has(idKey)) ids.add(obj.get(idKey).getAsString());
                }
            }
        } catch (Exception e) {
            System.err.println("❌ Error fetching IDs: " + e.getMessage());
        }
        return ids;
    }
    // ✅ POJO for TaskDetail
    static class TaskDetail {
        private String itemId;
        private String quantity;
        private String location;

        public TaskDetail(String itemId, String quantity, String location) {
            this.itemId = itemId;
            this.quantity = quantity;
            this.location = location;
        }

        public String getItemId() { return itemId; }
        public String getQuantity() { return quantity; }
        public String getLocation() { return location; }
    }


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

}