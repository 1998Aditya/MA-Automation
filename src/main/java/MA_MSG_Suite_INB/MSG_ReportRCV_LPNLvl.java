package MA_MSG_Suite_INB;

import org.apache.hc.client5.http.classic.HttpClient;
import org.apache.hc.client5.http.classic.methods.HttpPost;
import org.apache.hc.client5.http.entity.UrlEncodedFormEntity;
import org.apache.hc.client5.http.impl.classic.HttpClients;
import org.apache.hc.core5.http.ClassicHttpResponse;
import org.apache.hc.core5.http.NameValuePair;
import org.apache.hc.core5.http.message.BasicNameValuePair;
import org.apache.poi.ss.usermodel.*;
import org.json.JSONArray;
import org.json.JSONObject;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

public class MSG_ReportRCV_LPNLvl {

    private static final String EXCEL_PATH = "C://Users//Aditya Mishra//Desktop//Automation suite//auto_msg.xlsx";

    private static String USERNAME;
    private static String PASSWORD;
    private static String CLIENT_ID;
    private static String CLIENT_SECRET;
    private static String SELECTED_ORG;
    private static String SELECTED_LOC;
    private static String LOGIN_URL;
    private static String BASE_URL;
    private static String REPORT_URL;

    public void execute() {
        try {
            System.out.println("=== Step 4: ReportReceived iLPN API Started ===");
            readConfigFromExcel(EXCEL_PATH);
            String token = getAuthToken();
            if (token != null && !token.isEmpty()) {
                System.out.println("✅ Access Token retrieved successfully.\n");
                processReportFromExcel(token, EXCEL_PATH);
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
            REPORT_URL = BASE_URL + "device-integration/api/deviceintegration/process/ReportReceivediLPN_FER_Src_EP";

            System.out.println("Loaded Login Info:");
            System.out.println("Username=" + USERNAME + ", ClientId=" + CLIENT_ID);
            System.out.println("SelectedOrganization=" + SELECTED_ORG + ", SelectedLocation=" + SELECTED_LOC);
            System.out.println("Login URL=" + LOGIN_URL);
            System.out.println("Report URL=" + REPORT_URL + "\n");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // ✅ Token retrieval (same as ASN creation)
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

    // ✅ Build JSON from ReportRCV sheet and send to API
    private static void processReportFromExcel(String token, String filePath) {
        Workbook workbook = null;
        FileInputStream fis = null;

        try {
            System.out.println("📘 Reading 'ReportRCV' data from Excel...");
            File file = new File(filePath);
            fis = new FileInputStream(file);
            workbook = WorkbookFactory.create(fis);

            Sheet sheet = workbook.getSheet("ReportRCV");
            if (sheet == null) {
                System.out.println("❌ Sheet 'ReportRCV' not found. Check Excel.");
                return;
            }

            Iterator<Row> rowIterator = sheet.iterator();
            Row header = rowIterator.next(); // skip header
            DataFormatter formatter = new DataFormatter();

            // ✅ Find LCID column dynamically
            int lcidColumn = -1;
            for (int i = 0; i < header.getLastCellNum(); i++) {
                if ("Lcid".equalsIgnoreCase(getCellValue(header.getCell(i)))) {
                    lcidColumn = i;
                    break;
                }
            }
            if (lcidColumn == -1) {
                lcidColumn = header.getLastCellNum();
                header.createCell(lcidColumn).setCellValue("Lcid");
            }

            int totalRows = 0;
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                totalRows++;

                String asnId = getCellValue(row.getCell(0));
                String itemId = getCellValue(row.getCell(2));
                String LpnId = getCellValue(row.getCell(4));
                double qty = 0.0;
                try {
                    String qtyStr = formatter.formatCellValue(row.getCell(1)).trim();
                    if (!qtyStr.isEmpty()) qty = Double.parseDouble(qtyStr);
                } catch (Exception e) {
                    qty = 0.0;
                }

                if (asnId.isEmpty() || itemId.isEmpty()) {
                    System.out.println("⚠ Skipping blank row " + row.getRowNum());
                    continue;
                }

                // ✅ Call LPNgenerator for both LpnId and Lcid
                String lcid = "12345";//LPNgenerator.getNextLpn();

                // ✅ Write LCID to Excel
                Cell lcidCell = row.getCell(lcidColumn);
                if (lcidCell == null) lcidCell = row.createCell(lcidColumn);
                lcidCell.setCellValue(lcid);

                JSONObject detail = new JSONObject();
                detail.put("ItemId", itemId);
                detail.put("Quantity", qty);
                detail.put("PurchaseOrderId", "");

                JSONArray details = new JSONArray();
                details.put(detail);

                JSONObject body = new JSONObject();
                body.put("AsnId", asnId);
                body.put("Owner", "SHG"); // static for now
                body.put("LpnId", LpnId);
                body.put("Lcid", lcid);
                body.put("LpnDetail", details);
                body.put("MessageType", "ReportReceivediLPN");

                String uniqueKey = new SimpleDateFormat("yyyyMMddHHmmssSSS").format(new Date());
                body.put("UniqueKey", uniqueKey);

                System.out.println("\n========== JSON Payload for Row " + row.getRowNum() + " ==========");
                System.out.println(body.toString(4));
                System.out.println("=========================================================\n");

                //boolean success = triggerReportAPI(token, body, asnId);
                //System.out.println(success ? "✅ Report sent successfully for " + asnId : "❌ Report failed for " + asnId);
            }

            // ✅ Save LCIDs safely to Excel
            fis.close();
            try (FileOutputStream fos = new FileOutputStream(file)) {
                workbook.write(fos);
            }

            System.out.println("✅ LCID values written successfully. Processed total rows: " + totalRows);

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
        // ✅ Trigger ReportReceivediLPN API
        private static boolean triggerReportAPI(String token, JSONObject body, String asnId) {
            try {
                HttpClient client = HttpClients.createDefault();
                HttpPost post = new HttpPost(REPORT_URL);

                post.setHeader("Authorization", "Bearer " + token);
                post.setHeader("Content-Type", "application/json");
                post.setHeader("SelectedOrganization", SELECTED_ORG);
                post.setHeader("SelectedLocation", SELECTED_LOC);

                HttpEntity entity = new StringEntity(body.toString());
                post.setEntity(entity);

                ClassicHttpResponse response = (ClassicHttpResponse) client.execute(post);
                int statusCode = response.getCode();

                BufferedReader rd = new BufferedReader(new InputStreamReader(response.getEntity().getContent()));
                StringBuilder result = new StringBuilder();
                String line;
                while ((line = rd.readLine()) != null) result.append(line);

                System.out.println("Response for ASN " + asnId + " (HTTP " + statusCode + "): " + result);

                return (statusCode == 200 || statusCode == 201);

            } catch (Exception e) {
                e.printStackTrace();
                return false;
            }
        }
    */
    private static String getCellValue(Cell cell) {
        if (cell == null) return "";
        DataFormatter formatter = new DataFormatter();
        return formatter.formatCellValue(cell).trim();
    }
}
