package MA_MSG_Suite_INB;

import org.apache.hc.client5.http.classic.HttpClient;
import org.apache.hc.client5.http.classic.methods.HttpGet;
import org.apache.hc.client5.http.entity.UrlEncodedFormEntity;
import org.apache.hc.client5.http.impl.classic.HttpClients;
import org.apache.hc.client5.http.classic.methods.HttpPost;
import org.apache.hc.core5.http.ClassicHttpResponse;
import org.apache.hc.core5.http.NameValuePair;
import org.apache.hc.core5.http.message.BasicNameValuePair;
import org.apache.poi.ss.usermodel.*;
import org.json.JSONArray;
import org.json.JSONObject;

import java.io.*;
import java.util.*;

public class GetConditionCode {

    private static final String EXCEL_PATH = "C://Users//Aditya Mishra//Desktop//Automation suite//auto_msg.xlsx";

    private static String USERNAME;
    private static String PASSWORD;
    private static String CLIENT_ID;
    private static String CLIENT_SECRET;
    private static String SELECTED_ORG;
    private static String SELECTED_LOC;
    private static String LOGIN_URL;
    private static String BASE_URL;

    public void execute() {
        try {
            System.out.println("=== Step 5: Fetching Condition Codes for LCIDs ===");
            readConfigFromExcel(EXCEL_PATH);
            String token = getAuthToken();
            if (token != null && !token.isEmpty()) {
                System.out.println("✅ Access Token retrieved successfully.\n");
                fetchConditionCodes(token);
            } else {
                System.out.println("❌ Failed to retrieve token. Please verify credentials.");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // ✅ Read Login Configuration
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

            System.out.println("Loaded Login Info:");
            System.out.println("Username=" + USERNAME + ", ClientId=" + CLIENT_ID);
            System.out.println("SelectedOrganization=" + SELECTED_ORG + ", SelectedLocation=" + SELECTED_LOC);
            System.out.println("Login URL=" + LOGIN_URL);
            System.out.println("Base URL=" + BASE_URL + "\n");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // ✅ Authentication Token
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

            BufferedReader rd = new BufferedReader(new InputStreamReader(response.getEntity().getContent()));
            StringBuilder result = new StringBuilder();
            String line;
            while ((line = rd.readLine()) != null) result.append(line);

            int status = response.getCode();
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

    // ✅ Fetch and store condition codes for all LCIDs
    private static void fetchConditionCodes(String token) {
        Workbook workbook = null;
        FileInputStream fis = null;
        File file = new File(EXCEL_PATH);

        try {
            System.out.println("📘 Reading 'ReportRCV' sheet for LCIDs...");
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

            if (lcidColumn == -1) {
                System.out.println("❌ 'Lcid' column not found in sheet!");
                return;
            }
            if (conditionColumn == -1) {
                conditionColumn = header.getLastCellNum();
                header.createCell(conditionColumn).setCellValue("Condition_Code");
            }

            int count = 0;
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                String lcid = getCellValue(row.getCell(lcidColumn));
                if (lcid.isEmpty()) continue;

                String url = BASE_URL + "inventory-management/api/inventory-management/ilpnInquiry/conditionCodes/ilpnId?ilpnId=" + lcid;
                System.out.println("🔍 Fetching ConditionCode for LCID: " + lcid);

                String conditionCodes = getConditionCodesFromAPI(token, url);
                System.out.println("➡ Condition Codes: " + conditionCodes);

                Cell conditionCell = row.getCell(conditionColumn);
                if (conditionCell == null) conditionCell = row.createCell(conditionColumn);
                conditionCell.setCellValue(conditionCodes);

                // ✅ If condition code contains "FI" or "CR", trigger BoxDelivered only for this LCID
                if (conditionCodes.contains("\"FI\"") || conditionCodes.contains("\"CR\"")) {
                    System.out.println("🚚 Triggering BoxDelivered for LCID: " + lcid);

                    try {
                        // Backup sheet content
                        FileInputStream backupFis = new FileInputStream(file);
                        Workbook backupWb = WorkbookFactory.create(backupFis);
                        Sheet backupSheet = backupWb.getSheet("ReportRCV");
                        List<List<String>> allData = new ArrayList<>();

                        for (Row r : backupSheet) {
                            List<String> rowData = new ArrayList<>();
                            for (int i = 0; i < r.getLastCellNum(); i++) {
                                rowData.add(getCellValue(r.getCell(i)));
                            }
                            allData.add(rowData);
                        }
                        backupFis.close();

                        // Keep only the header + current LCID
                        Sheet filteredSheet = backupWb.getSheet("ReportRCV");
                        Iterator<Row> it = filteredSheet.iterator();
                        Row headerRow = it.next();
                        int lcidIndex = -1;
                        for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                            if ("Lcid".equalsIgnoreCase(getCellValue(headerRow.getCell(i)))) {
                                lcidIndex = i;
                                break;
                            }
                        }

                        int lastRow = filteredSheet.getLastRowNum();
                        for (int i = lastRow; i > 0; i--) {
                            Row r = filteredSheet.getRow(i);
                            if (r == null) continue;
                            String thisLcid = getCellValue(r.getCell(lcidIndex));
                            if (!thisLcid.equalsIgnoreCase(lcid)) {
                                filteredSheet.removeRow(r);
                            }
                        }

                        // Save filtered Excel
                        try (FileOutputStream fos = new FileOutputStream(file)) {
                            backupWb.write(fos);
                        }
                        backupWb.close();

                        // ✅ Run BoxDelivered
                        BoxDelivered boxDelivered = new BoxDelivered();
                        boxDelivered.execute();

                        // ✅ Restore original data
                        Workbook restoreWb = WorkbookFactory.create(true);
                        Sheet restoreSheet = restoreWb.createSheet("ReportRCV");
                        for (int i = 0; i < allData.size(); i++) {
                            Row newRow = restoreSheet.createRow(i);
                            List<String> dataRow = allData.get(i);
                            for (int j = 0; j < dataRow.size(); j++) {
                                newRow.createCell(j).setCellValue(dataRow.get(j));
                            }
                        }

                        try (FileOutputStream fosRestore = new FileOutputStream(file)) {
                            restoreWb.write(fosRestore);
                        }
                        restoreWb.close();

                        System.out.println("✅ BoxDelivered completed and Excel restored for LCID: " + lcid);
                    } catch (Exception ex) {
                        System.out.println("⚠ Failed to trigger BoxDelivered for LCID " + lcid + ": " + ex.getMessage());
                        ex.printStackTrace();
                    }
                }

                count++;
            }

            fis.close();
            try (FileOutputStream fos = new FileOutputStream(file)) {
                workbook.write(fos);
            }

            System.out.println("✅ Successfully updated Condition_Code for " + count + " LCIDs.");

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

    // ✅ Call API and extract condition codes
    private static String getConditionCodesFromAPI(String token, String url) {
        String conditionCodes = "";
        try {
            HttpClient client = HttpClients.createDefault();
            HttpGet get = new HttpGet(url);

            get.setHeader("Authorization", "Bearer " + token);
            get.setHeader("Content-Type", "application/json");
            get.setHeader("SelectedOrganization", SELECTED_ORG);
            get.setHeader("SelectedLocation", SELECTED_LOC);

            ClassicHttpResponse response = (ClassicHttpResponse) client.execute(get);
            int status = response.getCode();

            BufferedReader rd = new BufferedReader(new InputStreamReader(response.getEntity().getContent()));
            StringBuilder result = new StringBuilder();
            String line;
            while ((line = rd.readLine()) != null) result.append(line);

            if (status == 200) {
                JSONObject jsonData = new JSONObject(result.toString());
                if (jsonData.has("data") && jsonData.get("data") instanceof JSONArray) {
                    JSONArray dataArray = jsonData.getJSONArray("data");
                    List<String> codes = new ArrayList<>();

                    for (int i = 0; i < dataArray.length(); i++) {
                        JSONObject item = dataArray.getJSONObject(i);
                        String code = item.optString("ConditionCode", "");
                        if (!code.equalsIgnoreCase("PP") && !code.isEmpty()) {
                            codes.add("\"" + code + "\"");
                        }
                    }

                    conditionCodes = String.join(", ", codes);
                }
            } else {
                System.out.println("⚠ API returned status " + status + " for LCID: " + url);
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
        return conditionCodes;
    }

    private static String getCellValue(Cell cell) {
        if (cell == null) return "";
        DataFormatter formatter = new DataFormatter();
        return formatter.formatCellValue(cell).trim();
    }
}
