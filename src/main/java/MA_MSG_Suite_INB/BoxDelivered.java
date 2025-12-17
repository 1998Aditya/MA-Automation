package MA_MSG_Suite_INB;

import org.apache.hc.client5.http.classic.HttpClient;
import org.apache.hc.client5.http.classic.methods.HttpPost;
import org.apache.hc.client5.http.entity.UrlEncodedFormEntity;
import org.apache.hc.client5.http.impl.classic.HttpClients;
import org.apache.hc.core5.http.ClassicHttpResponse;
import org.apache.hc.core5.http.HttpEntity;
import org.apache.hc.core5.http.NameValuePair;
import org.apache.hc.core5.http.message.BasicNameValuePair;
import org.apache.hc.core5.http.io.entity.StringEntity;
import org.apache.poi.ss.usermodel.*;
import org.json.JSONArray;
import org.json.JSONObject;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

public class BoxDelivered {

    private static final String EXCEL_PATH = ExcelReaderIB.DATA_EXCEL_PATH;
    private static final String LOGIN_EXCEL_PATH = ExcelReaderIB.LOGIN_EXCEL_PATH;

    private static String USERNAME;
    private static String PASSWORD;
    private static String CLIENT_ID;
    private static String CLIENT_SECRET;
    private static String SELECTED_ORG;
    private static String SELECTED_LOC;
    private static String LOGIN_URL;
    private static String BASE_URL;
    private static String BOX_DELIVERED_URL;

    public void execute() {
        try {
            System.out.println("=== Step 6: BoxDeliveredToWMS API Started ===");
            readConfigFromExcel(LOGIN_EXCEL_PATH);
            String token = getAuthToken();
            if (token != null && !token.isEmpty()) {
                System.out.println("✅ Access Token retrieved successfully.\n");
                triggerBoxDelivered(token);
            } else {
                System.out.println("❌ Failed to retrieve token. Please verify credentials.");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // ✅ Read login/config info
    private static void readConfigFromExcel(String filePath) {
        try (FileInputStream fis = new FileInputStream(new File(filePath));
             Workbook workbook = WorkbookFactory.create(fis)) {

            Sheet sheet = workbook.getSheet("Login");
            if (sheet == null) throw new RuntimeException("No sheet named 'Login' found in Excel!");

            Row row = sheet.getRow(1);

            LOGIN_URL = getCellValue(row.getCell(0));
            BASE_URL = getCellValue(row.getCell(1));
            USERNAME = getCellValue(row.getCell(2));
            PASSWORD = getCellValue(row.getCell(3));
            CLIENT_ID = getCellValue(row.getCell(4));
            CLIENT_SECRET = getCellValue(row.getCell(5));
            SELECTED_ORG = getCellValue(row.getCell(6));
            SELECTED_LOC = getCellValue(row.getCell(7));

            if (!BASE_URL.endsWith("/")) BASE_URL += "/";
            BOX_DELIVERED_URL = BASE_URL + "device-integration/api/deviceintegration/process/BoxDeliveredToWMS_FER_Src_EP";

            System.out.println("Loaded Login Info:");
            System.out.println("Username=" + USERNAME + ", ClientId=" + CLIENT_ID);
            System.out.println("SelectedOrganization=" + SELECTED_ORG + ", SelectedLocation=" + SELECTED_LOC);
            System.out.println("Login URL=" + LOGIN_URL);
            System.out.println("Box Delivered URL=" + BOX_DELIVERED_URL + "\n");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // ✅ Token retrieval
    private static String getAuthToken() {
        String token = null;
        try {
            HttpClient client = HttpClients.createDefault();
            HttpPost post = new HttpPost(LOGIN_URL);

            List<NameValuePair> params = new ArrayList<>();
            params.add(new BasicNameValuePair("grant_type", "password"));
            params.add(new BasicNameValuePair("username", USERNAME));
            params.add(new BasicNameValuePair("password", PASSWORD));
            params.add(new BasicNameValuePair("client_id", CLIENT_ID));
            params.add(new BasicNameValuePair("client_secret", CLIENT_SECRET));

            post.setEntity(new UrlEncodedFormEntity(params));
            ClassicHttpResponse response = (ClassicHttpResponse) client.execute(post);
            int status = response.getCode();

            BufferedReader rd = new BufferedReader(new InputStreamReader(response.getEntity().getContent()));
            StringBuilder result = new StringBuilder();
            String line;
            while ((line = rd.readLine()) != null) result.append(line);

            if (status == 200 || status == 201) {
                JSONObject json = new JSONObject(result.toString());
                token = json.optString("access_token", "");
            } else {
                System.out.println("❌ Auth failed: " + result);
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
        return token;
    }

    // ✅ Trigger BoxDeliveredToWMS API
    private static void triggerBoxDelivered(String token) {
        Workbook workbook = null;
        FileInputStream fis = null;
        File file = new File(EXCEL_PATH);
        Random random = new Random();

        try {
            System.out.println("📘 Reading 'ReportRCV' sheet for LCID and Condition_Code...");
            fis = new FileInputStream(file);
            workbook = WorkbookFactory.create(fis);

            Sheet sheet = workbook.getSheet("ReportRCV");
            if (sheet == null) {
                System.out.println("❌ Sheet 'ReportRCV' not found!");
                return;
            }

            Iterator<Row> rowIterator = sheet.iterator();
            Row header = rowIterator.next(); // Skip header

            // ✅ Find LCID and Condition_Code columns dynamically and Testcase column
            int lcidColumn = -1;
            int conditionColumn = -1;
            int testcaseCol = -1;
            for (int i = 0; i < header.getLastCellNum(); i++) {
                String headerName = getCellValue(header.getCell(i));
                if ("Lcid".equalsIgnoreCase(headerName)) lcidColumn = i;
                if ("Condition_Code".equalsIgnoreCase(headerName)) conditionColumn = i;
                if ("Testcase".equalsIgnoreCase(headerName)) testcaseCol = i;
            }

            if (lcidColumn == -1) {
                System.out.println("❌ 'Lcid' column not found!");
                return;
            }

            boolean doFilterByTestcase = System.getProperty("testcase") != null && !System.getProperty("testcase").trim().isEmpty();
            String testcaseToRun = doFilterByTestcase ? System.getProperty("testcase").trim() : "";

            int count = 0;

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();

                // Testcase filter
                if (doFilterByTestcase) {
                    String tc = testcaseCol != -1 ? getCellValue(row.getCell(testcaseCol)) : getCellValue(row.getCell(0));
                    if (!testcaseToRun.equalsIgnoreCase(tc)) continue;
                }

                String lcid = getCellValue(row.getCell(lcidColumn));
                if (lcid.isEmpty()) continue;

                String conditionString = conditionColumn != -1 ? getCellValue(row.getCell(conditionColumn)) : "";
                List<String> conditionCodes = new ArrayList<>();
                if (!conditionString.isEmpty()) {
                    String[] parts = conditionString.replace("\"", "").split(",");
                    for (String code : parts) {
                        if (!code.trim().isEmpty()) {
                            conditionCodes.add(code.trim());
                        }
                    }
                }

                // ✅ Generate random station (MF_STATION_01 - MF_STATION_08)
                int stationNum = random.nextInt(8) + 1;
                String location = String.format("MF_STATION_%02d", stationNum);

                // ✅ Build ConditionCodes JSON
                JSONArray conditionArray = new JSONArray();
                for (String code : conditionCodes) {
                    JSONObject cond = new JSONObject();
                    cond.put("ConditionCode", code);
                    conditionArray.put(cond);
                }

                // ✅ Build LPN JSON
                JSONObject lpnObject = new JSONObject();
                lpnObject.put("LpnId", lcid);
                lpnObject.put("WCSOrderId", "");
                lpnObject.put("ConditionCodes", conditionArray);

                JSONArray lpnArray = new JSONArray();
                lpnArray.put(lpnObject);

                // ✅ Build full request body
                JSONObject body = new JSONObject();
                body.put("ToteId", "");
                body.put("Location", location);
                body.put("LPNs", lpnArray);
                body.put("MessageType", "BoxDeliveredToWMS");
                String uniqueKey = new SimpleDateFormat("yyyyMMddHHmmssSSS").format(new Date());
                body.put("UniqueKey", uniqueKey);

                System.out.println("\n========== JSON Payload for LCID " + lcid + " ==========");
                System.out.println(body.toString(4));
                System.out.println("=========================================================\n");

                boolean success = sendBoxDeliveredAPI(token, body, lcid);
                 System.out.println(success ? "✅ BoxDelivered sent successfully for " + lcid : "❌ BoxDelivered failed for " + lcid);
                count++;
            }

            System.out.println("✅ Total BoxDelivered messages sent: " + count);

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (fis != null) fis.close();
                if (workbook != null) workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }


        // ✅ Send API Request
        private static boolean sendBoxDeliveredAPI(String token, JSONObject body, String lcid) {
            try {
                HttpClient client = HttpClients.createDefault();
                HttpPost post = new HttpPost(BOX_DELIVERED_URL);

                post.setHeader("Authorization", "Bearer " + token);
                post.setHeader("Content-Type", "application/json");
                post.setHeader("SelectedOrganization", SELECTED_ORG);
                post.setHeader("SelectedLocation", SELECTED_LOC);

                HttpEntity entity = new StringEntity(body.toString());
                post.setEntity(entity);

                ClassicHttpResponse response = (ClassicHttpResponse) client.execute(post);
                int status = response.getCode();

                BufferedReader rd = new BufferedReader(new InputStreamReader(response.getEntity().getContent()));
                StringBuilder result = new StringBuilder();
                String line;
                while ((line = rd.readLine()) != null) result.append(line);

                System.out.println("Response for LCID " + lcid + " (HTTP " + status + "): " + result);
                return (status == 200 || status == 201);

            } catch (Exception e) {
                e.printStackTrace();
                return false;
            }
        }

    private static String getCellValue(Cell cell) {
        if (cell == null) return "";
        DataFormatter formatter = new DataFormatter();
        return formatter.formatCellValue(cell).trim();
    }
}
