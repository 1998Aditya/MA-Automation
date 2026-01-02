package MA_MSG_Suite_INB;

import org.apache.hc.client5.http.classic.HttpClient;
import org.apache.hc.client5.http.classic.methods.HttpPost;
import org.apache.hc.client5.http.entity.UrlEncodedFormEntity;        // client5
import org.apache.hc.client5.http.impl.classic.HttpClients;
import org.apache.hc.core5.http.ClassicHttpResponse;
import org.apache.hc.core5.http.NameValuePair;
import org.apache.hc.core5.http.io.entity.StringEntity;                // <-- correct package
import org.apache.hc.core5.http.message.BasicNameValuePair;
import org.apache.poi.ss.usermodel.*;
import org.json.JSONObject;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

public class iLPNToted {

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
    private static String TOTED_URL;

    public void execute() {
        try {
            System.out.println("=== Step: iLPNToted API Execution Started ===");
            readConfigFromExcel(LOGIN_EXCEL_PATH);
            String token = getAuthToken();

            if (token != null && token.length() > 0) {
                System.out.println("✅ Access Token retrieved successfully.\n");
                processTotedMessages(token);
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
            USERNAME = getCellValue(row.getCell(2));  // Used as UserId in payload
            PASSWORD = getCellValue(row.getCell(3));
            CLIENT_ID = getCellValue(row.getCell(4));
            CLIENT_SECRET = getCellValue(row.getCell(5));
            SELECTED_ORG = getCellValue(row.getCell(6));
            SELECTED_LOC = getCellValue(row.getCell(7));

            if (!BASE_URL.endsWith("/")) BASE_URL += "/";
            TOTED_URL = BASE_URL + "device-integration/api/deviceintegration/process/iLPNToted_FER_Src_EP";

            System.out.println("Loaded Login Info:");
            System.out.println("Username=" + USERNAME + ", ClientId=" + CLIENT_ID);
            System.out.println("SelectedOrganization=" + SELECTED_ORG + ", SelectedLocation=" + SELECTED_LOC);
            System.out.println("Login URL=" + LOGIN_URL);
            System.out.println("Toted URL=" + TOTED_URL + "\n");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // ✅ Authentication Token Retrieval
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

    // ✅ Process and send iLPNToted message for each LCID in ReportRCV sheet
    private static void processTotedMessages(String token) {
        Workbook workbook = null;
        FileInputStream fis = null;

        try {
            fis = new FileInputStream(EXCEL_PATH);
            workbook = WorkbookFactory.create(fis);

            Sheet sheet = workbook.getSheet("ReportRCV");
            if (sheet == null) {
                System.out.println("❌ Sheet 'ReportRCV' not found!");
                return;
            }

            Iterator<Row> rowIterator = sheet.iterator();
            if (rowIterator.hasNext()) rowIterator.next(); // Skip header row

            // Detect testcase column if present
            Row headerRow = sheet.getRow(0);
            int testcaseCol = -1;
            for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                if ("Testcase".equalsIgnoreCase(getCellValue(headerRow.getCell(i)))) {
                    testcaseCol = i;
                    break;
                }
            }
            boolean doFilterByTestcase = System.getProperty("testcase") != null && !System.getProperty("testcase").trim().isEmpty();
            String testcaseToRun = doFilterByTestcase ? System.getProperty("testcase").trim() : "";

            int count = 0;
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();

                // testcase filter
                if (doFilterByTestcase) {
                    String tc = testcaseCol != -1 ? getCellValue(row.getCell(testcaseCol)) : getCellValue(row.getCell(0));
                    if (!testcaseToRun.equalsIgnoreCase(tc)) continue;
                }

                String lcid = getCellValue(row.getCell(5)); // Assuming LCID is in column 5

                if (lcid == null || lcid.trim().isEmpty()) continue;

                JSONObject body = new JSONObject();
                body.put("LpnId", lcid);
                body.put("UserId", USERNAME);
                body.put("MessageType", "iLPNToted");

                String uniqueKey = new SimpleDateFormat("yyyyMMddHHmmssSSS").format(new Date());
                body.put("UniqueKey", uniqueKey);

                System.out.println("\n========== JSON Payload for LCID " + lcid + " ==========");
                System.out.println(body.toString(4));
                System.out.println("=========================================================\n");

                boolean success = triggerTotedAPI(token, body, lcid);
                 System.out.println(success ? "✅ iLPNToted sent successfully for LCID: " + lcid
                     : "❌ iLPNToted failed for LCID: " + lcid);

                count++;
            }

            System.out.println("\n✅ Total iLPNToted messages processed: " + count);

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

        // ✅ Trigger API for iLPNToted
    private static boolean triggerTotedAPI(String token, JSONObject body, String lcid) {
        try {
            HttpClient client = HttpClients.createDefault();
            HttpPost post = new HttpPost(TOTED_URL);

            post.setHeader("Authorization", "Bearer " + token);
            post.setHeader("Content-Type", "application/json");
            post.setHeader("SelectedOrganization", SELECTED_ORG);
            post.setHeader("SelectedLocation", SELECTED_LOC);

            post.setEntity(new StringEntity(body.toString()));
            ClassicHttpResponse response = (ClassicHttpResponse) client.execute(post);

            int statusCode = response.getCode();
            BufferedReader rd = new BufferedReader(new InputStreamReader(response.getEntity().getContent()));
            StringBuilder result = new StringBuilder();
            String line;
            while ((line = rd.readLine()) != null) result.append(line);

            System.out.println("Response for LCID " + lcid + " (HTTP " + statusCode + "): " + result);

            return (statusCode == 200 || statusCode == 201);

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
