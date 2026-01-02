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
import org.json.JSONObject;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

public class AsnInductComplete {

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
    private static String ASN_INDUCT_URL;

    public void execute() {
        try {
            System.out.println("=== Step: AsnInductComplete API Execution Started ===");

            readConfigFromExcel(LOGIN_EXCEL_PATH);
            String token = getAuthToken();

            if (token != null && !token.isEmpty()) {
                System.out.println("✅ Access Token retrieved successfully.\n");
                processAsnInduct(token);
            } else {
                System.out.println("❌ Failed to retrieve token.");
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // ✅ Read Login Configuration
    private static void readConfigFromExcel(String filePath) {
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = WorkbookFactory.create(fis)) {

            Sheet sheet = workbook.getSheet("Login");
            if (sheet == null)
                throw new RuntimeException("No sheet named 'Login' found!");

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
            ASN_INDUCT_URL = BASE_URL +
                    "device-integration/api/deviceintegration/process/AsnInductComplete_FER_Src_EP";

            System.out.println("Loaded Login Info:");
            System.out.println("Username=" + USERNAME);
            System.out.println("SelectedOrganization=" + SELECTED_ORG +
                    ", SelectedLocation=" + SELECTED_LOC);
            System.out.println("Login URL=" + LOGIN_URL);
            System.out.println("AsnInduct URL=" + ASN_INDUCT_URL + "\n");

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
            ClassicHttpResponse response =
                    (ClassicHttpResponse) client.execute(post);

            BufferedReader rd = new BufferedReader(
                    new InputStreamReader(response.getEntity().getContent()));

            StringBuilder result = new StringBuilder();
            String line;
            while ((line = rd.readLine()) != null) result.append(line);

            if (response.getCode() == 200 || response.getCode() == 201) {
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

    // ✅ Process ASN Induct Complete messages
    private static void processAsnInduct(String token) {
        try (FileInputStream fis = new FileInputStream(EXCEL_PATH);
             Workbook workbook = WorkbookFactory.create(fis)) {

            Sheet sheet = workbook.getSheet("ReportRCV");
            if (sheet == null) {
                System.out.println("❌ Sheet 'ReportRCV' not found!");
                return;
            }

            Row headerRow = sheet.getRow(0);
            int testcaseCol = -1;
            for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                if ("Testcase".equalsIgnoreCase(getCellValue(headerRow.getCell(i)))) {
                    testcaseCol = i;
                    break;
                }
            }

            boolean filterByTc = System.getProperty("testcase") != null &&
                    !System.getProperty("testcase").trim().isEmpty();
            String testcaseToRun = filterByTc
                    ? System.getProperty("testcase").trim()
                    : "";

            int count = 0;

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                if (filterByTc) {
                    String tc = testcaseCol != -1
                            ? getCellValue(row.getCell(testcaseCol))
                            : getCellValue(row.getCell(0));
                    if (!testcaseToRun.equalsIgnoreCase(tc)) continue;
                }

                String asnId = getCellValue(row.getCell(0));
                if (asnId.isEmpty()) continue;

                JSONObject body = new JSONObject();
                body.put("AsnId", asnId);
                body.put("Owner", "SHG");
                body.put("MessageType", "AsnInductComplete");
                body.put("UniqueKey",
                        new SimpleDateFormat("yyyyMMddHHmmssSSS")
                                .format(new Date()));

                System.out.println("\n========== ASN Payload ==========");
                System.out.println(body.toString(4));
                System.out.println("================================\n");

                boolean success = triggerAsnInductAPI(token, body, asnId);
                System.out.println(success
                        ? "✅ AsnInductComplete sent for ASN: " + asnId
                        : "❌ AsnInductComplete failed for ASN: " + asnId);

                count++;
            }

            System.out.println("\n✅ Total ASN Induct messages processed: " + count);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // ✅ Trigger API
    private static boolean triggerAsnInductAPI(String token,
                                               JSONObject body,
                                               String asnId) {
        try {
            HttpClient client = HttpClients.createDefault();
            HttpPost post = new HttpPost(ASN_INDUCT_URL);

            post.setHeader("Authorization", "Bearer " + token);
            post.setHeader("Content-Type", "application/json");
            post.setHeader("SelectedOrganization", SELECTED_ORG);
            post.setHeader("SelectedLocation", SELECTED_LOC);

            post.setEntity(new StringEntity(body.toString()));
            ClassicHttpResponse response =
                    (ClassicHttpResponse) client.execute(post);

            BufferedReader rd = new BufferedReader(
                    new InputStreamReader(response.getEntity().getContent()));

            StringBuilder result = new StringBuilder();
            String line;
            while ((line = rd.readLine()) != null) result.append(line);

            System.out.println("Response for ASN " + asnId +
                    " (HTTP " + response.getCode() + "): " + result);

            return response.getCode() == 200 || response.getCode() == 201;

        } catch (Exception e) {
            e.printStackTrace();
            return false;
        }
    }

    private static String getCellValue(Cell cell) {
        if (cell == null) return "";
        return new DataFormatter().formatCellValue(cell).trim();
    }
}
