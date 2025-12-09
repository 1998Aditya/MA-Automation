package MA_MSG_Suite_INB;

import org.apache.hc.client5.http.classic.HttpClient;
import org.apache.hc.client5.http.classic.methods.HttpPost;
import org.apache.hc.client5.http.entity.UrlEncodedFormEntity;
import org.apache.hc.client5.http.impl.classic.HttpClients;
import org.apache.hc.core5.http.ClassicHttpResponse;
import org.apache.hc.core5.http.NameValuePair;
import org.apache.hc.core5.http.io.entity.StringEntity;
import org.apache.hc.core5.http.message.BasicNameValuePair;
import org.apache.poi.ss.usermodel.*;
import org.json.JSONArray;
import org.json.JSONObject;

import java.io.*;
import java.util.*;

public class RemoveConditionCode {

    private static final String EXCEL_PATH = "C://Users//Aditya Mishra//Desktop//Automation suite//auto_msg.xlsx";

    private static String USERNAME;
    private static String PASSWORD;
    private static String CLIENT_ID;
    private static String CLIENT_SECRET;
    private static String SELECTED_ORG;
    private static String SELECTED_LOC;
    private static String LOGIN_URL;
    private static String BASE_URL;
    private static String REMOVE_CONDITION_URL;

    public void execute() {
        try {
            System.out.println("=== Step 7: Removing Condition Codes from LCIDs ===");
            readConfigFromExcel(EXCEL_PATH);
            String token = getAuthToken();
            if (token != null && !token.isEmpty()) {
                System.out.println("✅ Access Token retrieved successfully.\n");
                processRemoveConditionCodes(token);
            } else {
                System.out.println("❌ Failed to retrieve token. Please verify credentials.");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // ✅ Read Login Info
    private static void readConfigFromExcel(String filePath) {
        try (FileInputStream fis = new FileInputStream(new File(filePath));
             Workbook workbook = WorkbookFactory.create(fis)) {

            Sheet sheet = workbook.getSheet("Login");
            if (sheet == null)
                throw new RuntimeException("No sheet named 'Login' found in Excel!");

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
            REMOVE_CONDITION_URL = BASE_URL + "inventory-management/api/inventory-management/conditionAssignment/removeConditionCode";

            System.out.println("Loaded Login Info:");
            System.out.println("Username=" + USERNAME + ", ClientId=" + CLIENT_ID);
            System.out.println("SelectedOrganization=" + SELECTED_ORG + ", SelectedLocation=" + SELECTED_LOC);
            System.out.println("Login URL=" + LOGIN_URL);
            System.out.println("Remove Condition URL=" + REMOVE_CONDITION_URL + "\n");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // ✅ Generate Auth Token
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

    // ✅ Process LCIDs with existing Condition Codes
    private static void processRemoveConditionCodes(String token) {
        Workbook workbook = null;
        FileInputStream fis = null;
        File file = new File(EXCEL_PATH);

        try {
            System.out.println("📘 Reading 'ReportRCV' sheet for LCIDs with Condition Codes...");
            fis = new FileInputStream(file);
            workbook = WorkbookFactory.create(fis);
            Sheet sheet = workbook.getSheet("ReportRCV");

            if (sheet == null) {
                System.out.println("❌ Sheet 'ReportRCV' not found!");
                return;
            }

            Iterator<Row> rowIterator = sheet.iterator();
            Row header = rowIterator.next();

            int lcidColumn = -1;
            int conditionColumn = -1;

            for (int i = 0; i < header.getLastCellNum(); i++) {
                String headerName = getCellValue(header.getCell(i));
                if ("Lcid".equalsIgnoreCase(headerName)) lcidColumn = i;
                if ("Condition_Code".equalsIgnoreCase(headerName)) conditionColumn = i;
            }

            if (lcidColumn == -1 || conditionColumn == -1) {
                System.out.println("❌ Required columns not found in sheet!");
                return;
            }

            int processed = 0;
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                String lcid = getCellValue(row.getCell(lcidColumn));
                String conditionCodes = getCellValue(row.getCell(conditionColumn));

                if (lcid.isEmpty() || conditionCodes.isEmpty()) continue;

                // Parse condition codes
                List<String> codeList = new ArrayList<>();
                String cleaned = conditionCodes.replace("\"", "").trim();
                if (!cleaned.isEmpty()) {
                    for (String code : cleaned.split(",")) {
                        if (!code.trim().isEmpty()) codeList.add(code.trim());
                    }
                }

                if (codeList.isEmpty()) continue;

                // Prepare JSON Body
                JSONObject body = new JSONObject();
                JSONArray codeArray = new JSONArray();
                for (String c : codeList) codeArray.put(c);

                body.put("conditionCodeList", codeArray);
                body.put("containerId", lcid);
                body.put("containerType", "ILPN");
                body.put("criteriaId", "Basic Condition Code UnAssignment");
                body.put("transactionId", "Condition Code UnAssignment");
                body.put("validateIlpnOnly", true);

                System.out.println("\n========== JSON Payload for LCID " + lcid + " ==========");
                System.out.println(body.toString(4));
                System.out.println("=========================================================\n");

             /*   // Trigger API
                boolean success = callRemoveConditionAPI(token, body, lcid);
                if (success) {
                    System.out.println("✅ Condition Code(s) removed for LCID: " + lcid);
                } else {
                    System.out.println("❌ Failed to remove Condition Code for LCID: " + lcid);
                }*/
                processed++;
            }

            System.out.println("✅ Process completed for " + processed + " LCIDs.");

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
/*
    // ✅ Trigger Remove Condition Code API
    private static boolean callRemoveConditionAPI(String token, JSONObject body, String lcid) {
        try {
            HttpClient client = HttpClients.createDefault();
            HttpPost post = new HttpPost(REMOVE_CONDITION_URL);

            post.setHeader("Authorization", "Bearer " + token);
            post.setHeader("Content-Type", "application/json");
            post.setHeader("SelectedOrganization", SELECTED_ORG);
            post.setHeader("SelectedLocation", SELECTED_LOC);

            post.setEntity(new StringEntity(body.toString()));

            ClassicHttpResponse response = (ClassicHttpResponse) client.execute(post);
            int status = response.getCode();

            BufferedReader rd = new BufferedReader(new InputStreamReader(response.getEntity().getContent()));
            StringBuilder result = new StringBuilder();
            String line;
            while ((line = rd.readLine()) != null) result.append(line);

            System.out.println("Response for LCID " + lcid + " (HTTP " + status + "): " + result);
            return (status == 200 || status == 201);

        } catch (Exception e) {
            System.out.println("⚠ Exception while removing ConditionCode for " + lcid + ": " + e.getMessage());
            return false;
        }
    }
*/
    // ✅ Utility for reading cell values
    private static String getCellValue(Cell cell) {
        if (cell == null) return "";
        DataFormatter formatter = new DataFormatter();
        return formatter.formatCellValue(cell).trim();
    }
}
