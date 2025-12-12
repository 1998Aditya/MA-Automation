package MA_MSG_Suite_INB;

import org.apache.hc.client5.http.classic.methods.HttpPost;
import org.apache.hc.client5.http.classic.methods.HttpGet;
import org.apache.hc.client5.http.classic.HttpClient;
import org.apache.hc.client5.http.impl.classic.HttpClients;
import org.apache.hc.client5.http.entity.UrlEncodedFormEntity;
import org.apache.hc.core5.http.ClassicHttpResponse;
import org.apache.hc.core5.http.HttpEntity;
import org.apache.hc.core5.http.NameValuePair;
import org.apache.hc.core5.http.message.BasicNameValuePair;
import org.apache.hc.core5.http.io.entity.StringEntity;

import java.io.*;
import java.util.*;

import org.apache.poi.ss.usermodel.*;
import org.json.JSONArray;
import org.json.JSONObject;

public class MSG_Item_ASN_Creation {

    // ✅ updated path for MSG
    private static final String EXCEL_PATH = "C://Users//Aditya Mishra//Desktop//Automation suite//auto_msg.xlsx";

    private static String USERNAME;
    private static String PASSWORD;
    private static String CLIENT_ID;
    private static String CLIENT_SECRET;
    private static String SELECTED_ORG;
    private static String SELECTED_LOC;
    private static String LOGIN_URL;
    private static String BASE_URL;
    private static String TRIGGER_URL;
    private static String CHECK_URL;

    // ✅ Called from MSG_MAIN
    public void execute() {
        try {
            System.out.println("=== Step 3: ASN Creation Started ===");
            readConfigFromExcel(EXCEL_PATH);
            String token = getAuthToken();
            if (token != null && !token.isEmpty()) {
                System.out.println("Access Token retrieved successfully.\n");
                createASNsFromExcel(token, EXCEL_PATH);
            } else {
                System.out.println("❌ Failed to retrieve access token. Please verify credentials.");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

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

            if (!BASE_URL.endsWith("/")) BASE_URL += "/"; // ensure correct concatenation

            TRIGGER_URL = BASE_URL + "receiving/api/receiving/asn/bulkImport";
            CHECK_URL = BASE_URL + "receiving/api/receiving/asn/asnId/";

            System.out.println("Loaded Login Info:");
            System.out.println("Username=" + USERNAME + ", ClientId=" + CLIENT_ID);
            System.out.println("SelectedOrganization=" + SELECTED_ORG + ", SelectedLocation=" + SELECTED_LOC);
            System.out.println("Login URL=" + LOGIN_URL);
            System.out.println("Trigger URL=" + TRIGGER_URL);
            System.out.println("Check URL=" + CHECK_URL + "\n");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

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
            System.out.println("Auth Response Code: " + status);

            BufferedReader rd = new BufferedReader(new InputStreamReader(response.getEntity().getContent()));
            StringBuilder result = new StringBuilder();
            String line;
            while ((line = rd.readLine()) != null) result.append(line);

            System.out.println("Auth Response Body: " + result);

            if (status == 200 || status == 201) {
                JSONObject json = new JSONObject(result.toString());
                token = json.optString("access_token", "");
            } else {
                System.out.println("❌ Authentication failed: " + result);
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
        return token;
    }

    private static boolean checkASNExists(String token, String asnId) {
        try {
            HttpClient client = HttpClients.createDefault();
            HttpGet get = new HttpGet(CHECK_URL + asnId);

            get.setHeader("Authorization", "Bearer " + token);
            get.setHeader("SelectedOrganization", SELECTED_ORG);
            get.setHeader("SelectedLocation", SELECTED_LOC);

            ClassicHttpResponse response = (ClassicHttpResponse) client.execute(get);
            int statusCode = response.getCode();

            if (statusCode == 200) {
                System.out.println("ASN " + asnId + " already exists. Skipping...");
                return true;
            } else if (statusCode == 404) {
                System.out.println("ASN " + asnId + " does not exist. Proceeding to create...");
                return false;
            } else {
                System.out.println("Unexpected response while checking ASN " + asnId + ": " + statusCode);
                return false;
            }
        } catch (Exception e) {
            e.printStackTrace();
            return false;
        }
    }

    private static void createASNsFromExcel(String token, String filePath) {
        Workbook workbook = null;
        try (FileInputStream fis = new FileInputStream(new File(filePath))) {

            workbook = WorkbookFactory.create(fis);
            Sheet sheet = workbook.getSheet("Item_ASN");
            if (sheet == null) throw new RuntimeException("No sheet named 'Item_ASN' found in Excel!");

            Iterator<Row> rowIterator = sheet.iterator();
            Row header = rowIterator.next(); // skip header row

            // Find Run_status column index
            int runStatusCol = -1;
            for (int i = 0; i < header.getLastCellNum(); i++) {
                if ("Run_status".equalsIgnoreCase(header.getCell(i).getStringCellValue().trim())) {
                    runStatusCol = i;
                    break;
                }
            }
            if (runStatusCol == -1) throw new RuntimeException("No 'Run_status' column found in Excel!");

            DataFormatter formatter = new DataFormatter();
            Map<String, JSONObject> asnMap = new LinkedHashMap<>();
            Map<String, List<Row>> asnRowsMap = new HashMap<>();

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                String runStatus = getCellValue(row.getCell(runStatusCol));
                if ("Success".equalsIgnoreCase(runStatus)) continue;

                String asnId = getCellValue(row.getCell(0));
                if (asnId.isEmpty()) continue;

                asnRowsMap.computeIfAbsent(asnId, k -> new ArrayList<>()).add(row);

                String asnOriginTypeId = getCellValue(row.getCell(1));
                String destFacilityId = getCellValue(row.getCell(2));
                String destFacilityAliasId = getCellValue(row.getCell(3));
                String maujdsHostId = getCellValue(row.getCell(4));
                String maujdsIsMarked = getCellValue(row.getCell(5));
                String maujdsBookingRef = getCellValue(row.getCell(6));
                String maujdsBookingDate = getCellValue(row.getCell(7));
                String asnLineId = formatter.formatCellValue(row.getCell(8));

                // ✅ FIXED: Safe numeric parsing for shippedQty
                double shippedQty = 0.0;
                try {
                    String qtyStr = formatter.formatCellValue(row.getCell(9)).trim();
                    if (!qtyStr.isEmpty()) shippedQty = Double.parseDouble(qtyStr);
                } catch (Exception e) {
                    shippedQty = 0.0;
                }

                String itemId = getCellValue(row.getCell(10));
                String purchaseOrderId = getCellValue(row.getCell(11));
                String qtyUomId = getCellValue(row.getCell(12));

                JSONObject asnLineExtended = new JSONObject();
                asnLineExtended.put("MAUJDSBookingRef", maujdsBookingRef);
                asnLineExtended.put("MAUJDSBookingDate", maujdsBookingDate);

                JSONObject asnLine = new JSONObject();
                asnLine.put("Extended", asnLineExtended);
                asnLine.put("AsnLineId", asnLineId);
                asnLine.put("ShippedQuantity", shippedQty);
                asnLine.put("AsnId", asnId);
                asnLine.put("ItemId", itemId);
                asnLine.put("PurchaseOrderId", purchaseOrderId);
                asnLine.put("QuantityUomId", qtyUomId);

                if (asnMap.containsKey(asnId)) {
                    asnMap.get(asnId).getJSONArray("AsnLine").put(asnLine);
                } else {
                    JSONObject extended = new JSONObject();
                    extended.put("MAUJDSHostId", maujdsHostId);
                    extended.put("MAUJDSBookingRef", maujdsBookingRef);
                    extended.put("MAUJDSIsMarked", Boolean.parseBoolean(maujdsIsMarked));

                    JSONArray asnLines = new JSONArray();
                    asnLines.put(asnLine);

                    JSONObject asnObject = new JSONObject();
                    asnObject.put("actions", new JSONObject());
                    asnObject.put("AsnId", asnId);
                    asnObject.put("AsnOriginTypeId", asnOriginTypeId);
                    asnObject.put("DestinationFacilityId", destFacilityId);
                    asnObject.put("DestinationFacilityAliasId", destFacilityAliasId);
                    asnObject.put("Extended", extended);
                    asnObject.put("AsnLine", asnLines);

                    asnMap.put(asnId, asnObject);
                }
            }

            CellStyle successStyle = workbook.createCellStyle();
            successStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
            successStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            CellStyle failStyle = workbook.createCellStyle();
            failStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
            failStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            for (String asnId : asnMap.keySet()) {
                List<Row> rows = asnRowsMap.get(asnId);

                JSONObject body = new JSONObject();
                JSONArray dataArray = new JSONArray();
                dataArray.put(asnMap.get(asnId));
                body.put("Data", dataArray);
            /*
                // 👇 Print payload for debugging
                System.out.println("\n========== JSON Payload for ASN " + asnId + " ==========");
                System.out.println(body.toString(4));
                System.out.println("=========================================================\n");
            */
                boolean success = triggerASNAPI(token, body, asnId);

                for (Row r : rows) {
                    Cell statusCell = r.getCell(runStatusCol);
                    if (statusCell == null) statusCell = r.createCell(runStatusCol);
                    if (success) {
                        statusCell.setCellValue("Success");
                        statusCell.setCellStyle(successStyle);
                    } else {
                        statusCell.setCellValue("Failed");
                        statusCell.setCellStyle(failStyle);
                    }
                }
            }

            try (FileOutputStream fos = new FileOutputStream(filePath)) {
                workbook.write(fos);
            }

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (workbook != null) workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    private static boolean triggerASNAPI(String token, JSONObject body, String asnId) {
        try {
            HttpClient client = HttpClients.createDefault();
            HttpPost post = new HttpPost(TRIGGER_URL);

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

            if (statusCode == 200 || statusCode == 201) {
                JSONObject jsonResponse = new JSONObject(result.toString());
                return !jsonResponse.has("success") || jsonResponse.getBoolean("success");
            } else {
                System.out.println("❌ API call failed for ASN " + asnId + " with status " + statusCode);
                return false;
            }

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
