package MA_MSG_Suite_INB;

import org.apache.hc.client5.http.classic.HttpClient;
import org.apache.hc.client5.http.classic.methods.HttpGet;
import org.apache.hc.client5.http.classic.methods.HttpPost;
import org.apache.hc.client5.http.entity.UrlEncodedFormEntity;
import org.apache.hc.client5.http.impl.classic.HttpClients;
import org.apache.hc.core5.http.ClassicHttpResponse;
import org.apache.hc.core5.http.NameValuePair;
import org.apache.hc.core5.http.io.entity.StringEntity;
import org.apache.hc.core5.http.message.BasicNameValuePair;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;

import java.io.*;
import java.util.*;
import java.util.regex.Pattern;

public class MSG_LPN_ASN_Creation {

    // Existing ASN data excel
    //    private static final String EXCEL_PATH = "C://Users//Aditya Mishra//IdeaProjects//msg-runner//auto_msg.xlsx";
    private static final String EXCEL_PATH = ExcelReaderIB.DATA_EXCEL_PATH;
    // 👉 Login.xlsx only for Login sheet
    //    private static final String LOGIN_EXCEL_PATH = "C://Users//Aditya Mishra//IdeaProjects//msg-runner//Login.xlsx";
    private static final String LOGIN_EXCEL_PATH = ExcelReaderIB.LOGIN_EXCEL_PATH;

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

    // ✅ Called from MSG controller
    public void execute() {
        try {
            System.out.println("=== LPN ASN Creation Started ===");
            System.out.println("Reading Login config from: " + LOGIN_EXCEL_PATH);
            // 🔹 Read Login sheet ONLY from Login.xlsx
            readConfigFromExcel(LOGIN_EXCEL_PATH);

            String token = getAuthToken();
            System.out.println("Token after auth: " + (token == null ? "null" : (token.isEmpty() ? "EMPTY" : "NON-EMPTY")));
            if (token != null) {
                System.out.println("✅ Access Token Retrieved.");
                System.out.println("Calling createLPNASNsFromExcel with file: " + EXCEL_PATH);
                // 🔹 LPN_ASN data from auto.xlsx
                createLPNASNsFromExcel(token, EXCEL_PATH);
            } else {
                System.out.println("❌ Failed to retrieve token.");
            }
        } catch (Exception e) {
            System.out.println("❌ Exception in MSG_LPN_ASN_Creation.execute():");
            e.printStackTrace();
        }
    }

    private static void readConfigFromExcel(String filePath) {
        try (FileInputStream fis = new FileInputStream(new File(filePath));
             Workbook workbook = WorkbookFactory.create(fis)) {

            System.out.println("Opened Login workbook: " + filePath);

            Sheet sheet = workbook.getSheet("Login");
            if (sheet == null) throw new RuntimeException("No sheet named 'Login' found in Excel!");

            Row row = sheet.getRow(1);
            if (row == null) {
                throw new RuntimeException("Row 1 in Login sheet is null!");
            }

            LOGIN_URL = getCellValue(row.getCell(0));
            BASE_URL = getCellValue(row.getCell(1));
            USERNAME = getCellValue(row.getCell(2));
            PASSWORD = getCellValue(row.getCell(3));
            CLIENT_ID = getCellValue(row.getCell(4));
            CLIENT_SECRET = getCellValue(row.getCell(5));
            SELECTED_ORG = getCellValue(row.getCell(6));
            SELECTED_LOC = getCellValue(row.getCell(7));

            if (!BASE_URL.endsWith("/")) BASE_URL += "/";

            TRIGGER_URL = BASE_URL + "receiving/api/receiving/asn/bulkImport";
            CHECK_URL = BASE_URL + "receiving/api/receiving/asn/asnId/";

            System.out.println("✅ Loaded Login Info: " + USERNAME + " / Org=" + SELECTED_ORG + " / Loc=" + SELECTED_LOC);
            System.out.println("  Login URL   = " + LOGIN_URL);
            System.out.println("  Base URL    = " + BASE_URL);
            System.out.println("  Trigger URL = " + TRIGGER_URL);
            System.out.println("  Check URL   = " + CHECK_URL);

        } catch (Exception e) {
            System.out.println("❌ Exception in readConfigFromExcel():");
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

            JSONObject json = new JSONObject(result.toString());
            token = json.getString("access_token");

        } catch (Exception e) {
            System.out.println("❌ Exception in getAuthToken():");
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
                System.out.println("⚠️ ASN " + asnId + " already exists. Skipping...");
                return true;
            } else if (statusCode == 404) {
                System.out.println("✅ ASN " + asnId + " does not exist. Proceeding...");
                return false;
            } else {
                System.out.println("❌ Unexpected response for ASN " + asnId + ": " + statusCode);
                return true;
            }
        } catch (Exception e) {
            System.out.println("❌ Exception in checkASNExists():");
            e.printStackTrace();
            return true;
        }
    }

    /**
     * Core creation logic.
     * - Groups rows by Testcase (if Testcase column exists).
     * - Processes each testcase in order (LinkedHashMap preserves sheet order).
     * - Saves workbook after finishing each testcase block (safer).
     */
    private static void createLPNASNsFromExcel(String token, String filePath) {
        System.out.println("➡ Entering createLPNASNsFromExcel(), filePath = " + filePath);
        Workbook workbook = null;
        try (FileInputStream fis = new FileInputStream(new File(filePath))) {

            workbook = WorkbookFactory.create(fis);
            System.out.println("Opened ASN workbook: " + filePath);

            Sheet sheet = workbook.getSheet("LPN_ASN");
            if (sheet == null) throw new RuntimeException("No sheet named 'LPN_ASN' found in Excel!");

            System.out.println("LPN_ASN sheet lastRowNum = " + sheet.getLastRowNum());

            DataFormatter formatter = new DataFormatter();
            Iterator<Row> rowIterator = sheet.iterator();
            if (!rowIterator.hasNext()) {
                System.out.println("LPN_ASN sheet has no rows!");
                return;
            }

            Row header = rowIterator.next(); // skip header row
            System.out.println("Header row lastCellNum = " + header.getLastCellNum());

            // Find Run_status column index and Testcase column index
            int runStatusCol = -1;
            int testcaseCol = -1;
            for (int i = 0; i < header.getLastCellNum(); i++) {
                Cell c = header.getCell(i);
                if (c == null) continue;
                String val = getCellValue(c);
                System.out.println("Header col " + i + " = '" + val + "'");
                if ("Run_status".equalsIgnoreCase(val.trim())) {
                    runStatusCol = i;
                }
                if ("Testcase".equalsIgnoreCase(val.trim())) {
                    testcaseCol = i;
                }
            }
            System.out.println("Detected Run_status column index = " + runStatusCol);
            System.out.println("Detected Testcase column index = " + testcaseCol);
            if (runStatusCol == -1) throw new RuntimeException("No 'Run_status' column found in Excel!");

            // Group rows by Testcase (if present), else single bucket with key "_ALL"
            Map<String, List<Row>> testCaseMap;
            if (testcaseCol != -1) {
                testCaseMap = groupRowsByTestcasePreserveOrder(sheet, testcaseCol);
                if (testCaseMap.isEmpty()) {
                    System.out.println("⚠ No rows matched Testcase pattern; falling back to single-batch processing.");
                    testCaseMap = new LinkedHashMap<>();
                    // collect all non-header rows
                    Iterator<Row> it = sheet.iterator();
                    if (it.hasNext()) it.next();
                    while (it.hasNext()) {
                        Row r = it.next();
                        if (r == null) continue;
                        testCaseMap.computeIfAbsent("_ALL", k -> new ArrayList<>()).add(r);
                    }
                }
            } else {
                System.out.println("⚠ No 'Testcase' column found; processing entire sheet as single batch.");
                testCaseMap = new LinkedHashMap<>();
                Iterator<Row> it = sheet.iterator();
                if (it.hasNext()) it.next();
                while (it.hasNext()) {
                    Row r = it.next();
                    if (r == null) continue;
                    testCaseMap.computeIfAbsent("_ALL", k -> new ArrayList<>()).add(r);
                }
            }

            System.out.println("Testcase groups found: " + testCaseMap.keySet());

            // Styles for success/failure
            CellStyle successStyle = workbook.createCellStyle();
            successStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
            successStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            CellStyle failStyle = workbook.createCellStyle();
            failStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
            failStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            // For each testcase (in order) process only those rows
            for (String tcKey : testCaseMap.keySet()) {
                System.out.println("\n--- Processing Testcase: " + tcKey + " ---");
                List<Row> rowsForTc = testCaseMap.get(tcKey);

                // Build ASN map for this testcase
                Map<String, JSONObject> asnMap = new LinkedHashMap<>();
                Map<String, List<Row>> asnRowsMap = new HashMap<>();

                int processedRows = 0;
                int skippedRunStatus = 0;
                int skippedEmptyAsn = 0;

                for (Row row : rowsForTc) {
                    if (row == null) continue;
                    processedRows++;

                    String runStatus = getCellValue(row.getCell(runStatusCol));
                    String asnId = formatter.formatCellValue(row.getCell(0));

                    System.out.println("Row " + row.getRowNum() + " -> Run_status='" + runStatus + "', AsnId='" + asnId + "'");

                    if ("Success".equalsIgnoreCase(runStatus)) {
                        skippedRunStatus++;
                        continue; // skip successful rows
                    }

                    if (asnId == null || asnId.isEmpty()) {
                        skippedEmptyAsn++;
                        continue;
                    }

                    // Track all rows with same ASN
                    asnRowsMap.computeIfAbsent(asnId, k -> new ArrayList<>()).add(row);

                    String asnOriginTypeId = formatter.formatCellValue(row.getCell(1));
                    String vendorId = formatter.formatCellValue(row.getCell(2));
                    String destFacilityId = formatter.formatCellValue(row.getCell(3));
                    String maujdsHostId = formatter.formatCellValue(row.getCell(4));
                    String maujdsIsMarked = formatter.formatCellValue(row.getCell(5));
                    String maujdsASNToVNA = formatter.formatCellValue(row.getCell(6));
                    String maujdsIBTType = formatter.formatCellValue(row.getCell(7));

                    String lpnId = formatter.formatCellValue(row.getCell(8));
                    String shippedQty = formatter.formatCellValue(row.getCell(9));
                    String itemId = formatter.formatCellValue(row.getCell(10));
                    String purchaseOrderId = formatter.formatCellValue(row.getCell(11));
                    String purchaseOrderLineId = formatter.formatCellValue(row.getCell(12));
                    String qtyUomId = formatter.formatCellValue(row.getCell(13));

                    // Build LpnDetail
                    JSONObject lpnDetail = new JSONObject();
                    lpnDetail.put("LpnId", lpnId);
                    lpnDetail.put("PurchaseOrderId", purchaseOrderId.isEmpty() ? JSONObject.NULL : purchaseOrderId);
                    lpnDetail.put("PurchaseOrderLineId", purchaseOrderLineId.isEmpty() ? JSONObject.NULL : purchaseOrderLineId);
                    lpnDetail.put("ItemId", itemId);
                    lpnDetail.put("QuantityUomId", qtyUomId);
                    lpnDetail.put("ShippedQuantity", shippedQty);

                    JSONArray lpnDetailArray = new JSONArray();
                    lpnDetailArray.put(lpnDetail);

                    JSONObject lpn = new JSONObject();
                    lpn.put("AsnId", asnId);
                    lpn.put("LpnId", lpnId);
                    lpn.put("LpnDetail", lpnDetailArray);

                    // If ASN already exists in map → add Lpn
                    if (asnMap.containsKey(asnId)) {
                        asnMap.get(asnId).getJSONArray("Lpn").put(lpn);
                    } else {
                        JSONObject extended = new JSONObject();
                        extended.put("MAUJDSHostId", maujdsHostId.isEmpty() ? JSONObject.NULL : maujdsHostId);
                        extended.put("MAUJDSIsMarked", Boolean.parseBoolean(maujdsIsMarked));
                        extended.put("MAUJDSASNToVNA", Boolean.parseBoolean(maujdsASNToVNA));
                        extended.put("MAUJDSIBTType", maujdsIBTType.isEmpty() ? JSONObject.NULL : maujdsIBTType);

                        JSONArray lpnArray = new JSONArray();
                        lpnArray.put(lpn);

                        JSONObject asnObject = new JSONObject();
                        asnObject.put("Actions", new JSONObject());
                        asnObject.put("AsnId", asnId);
                        asnObject.put("AsnOriginTypeId", asnOriginTypeId);
                        asnObject.put("VendorId", vendorId.isEmpty() ? JSONObject.NULL : vendorId);
                        asnObject.put("DestinationFacilityId", destFacilityId);
                        asnObject.put("Extended", extended);
                        asnObject.put("Lpn", lpnArray);

                        asnMap.put(asnId, asnObject);
                    }
                }

                System.out.println("Finished building payloads for testcase " + tcKey);
                System.out.println("  Processed rows in this testcase  = " + processedRows);
                System.out.println("  Skipped (Run_status=Success)     = " + skippedRunStatus);
                System.out.println("  Skipped (empty AsnId)            = " + skippedEmptyAsn);
                System.out.println("  Unique ASN IDs in this testcase  = " + asnMap.size());

                if (asnMap.isEmpty()) {
                    System.out.println("⚠ No LPN ASN records found to send for testcase " + tcKey);
                }

                // Loop through all ASNs and post if not existing (only for this testcase)
                for (String asnId : asnMap.keySet()) {
                    List<Row> rows = asnRowsMap.get(asnId);

                    try {
                        if (checkASNExists(token, asnId)) {
                            for (Row r : rows) {
                                Cell statusCell = r.getCell(runStatusCol);
                                if (statusCell == null) statusCell = r.createCell(runStatusCol);
                                statusCell.setCellValue("Success");
                                statusCell.setCellStyle(successStyle);
                            }
                            continue;
                        }

                        JSONObject body = new JSONObject();
                        JSONArray dataArray = new JSONArray();
                        dataArray.put(asnMap.get(asnId));
                        body.put("Data", dataArray);

                        // 👇 Print payload for debugging
                        System.out.println("\n========== JSON Payload for LPN ASN " + asnId + " (Testcase " + tcKey + ") ==========");
                        System.out.println(body.toString(4));
                        System.out.println("============================================================\n");

                        boolean success = triggerLPNASNAPI(body, token, asnId);

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

                    } catch (Exception ex) {
                        System.out.println("❌ Exception while processing ASN " + asnId + ":");
                        ex.printStackTrace();
                        for (Row r : rows) {
                            Cell statusCell = r.getCell(runStatusCol);
                            if (statusCell == null) statusCell = r.createCell(runStatusCol);
                            statusCell.setCellValue("Failed");
                            statusCell.setCellStyle(failStyle);
                        }
                    }
                }

                // Save updates back to file after finishing this testcase batch
                try (FileOutputStream fos = new FileOutputStream(filePath)) {
                    workbook.write(fos);
                    System.out.println("✅ LPN ASN Excel updated and saved for testcase " + tcKey + ": " + filePath);
                } catch (Exception saveEx) {
                    System.out.println("❌ Failed to save workbook after testcase " + tcKey + ":");
                    saveEx.printStackTrace();
                }
            }

            System.out.println("\nAll testcases processed in LPN_ASN sheet.");

        } catch (Exception e) {
            System.out.println("❌ Exception in createLPNASNsFromExcel():");
            e.printStackTrace();
        } finally {
            try {
                if (workbook != null) workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    // Groups rows by Testcase column value (expects values like TST_1, TST_2, ...)
    // Preserves insertion/order using LinkedHashMap
    private static Map<String, List<Row>> groupRowsByTestcasePreserveOrder(Sheet sheet, int testcaseColIndex) {
        Map<String, List<Row>> map = new LinkedHashMap<>();
        if (sheet == null) return map;

        Iterator<Row> it = sheet.iterator();
        if (it.hasNext()) it.next(); // skip header

        Pattern p = Pattern.compile("^TST_\\d+$", Pattern.CASE_INSENSITIVE);

        while (it.hasNext()) {
            Row row = it.next();
            if (row == null) continue;
            Cell c = row.getCell(testcaseColIndex);
            if (c == null) continue;
            String tc = getCellValue(c).trim();
            if (!p.matcher(tc).matches()) continue;
            map.computeIfAbsent(tc, x -> new ArrayList<>()).add(row);
        }
        return map;
    }

    private static boolean triggerLPNASNAPI(JSONObject body, String token, String asnId) {
        try {
            HttpClient client = HttpClients.createDefault();
            HttpPost post = new HttpPost(TRIGGER_URL);

            post.setHeader("Authorization", "Bearer " + token);
            post.setHeader("SelectedOrganization", SELECTED_ORG);
            post.setHeader("SelectedLocation", SELECTED_LOC);
            post.setHeader("Content-Type", "application/json");

            post.setEntity(new StringEntity(body.toString()));

            // Print JSON being sent
            //            System.out.println("------ JSON Request for ASN " + asnId + " ------");
            //            System.out.println(body.toString(4));
            //            System.out.println("------------------------------------------------");

            ClassicHttpResponse response = (ClassicHttpResponse) client.execute(post);

            int statusCode = response.getCode();
            BufferedReader rd = new BufferedReader(new InputStreamReader(response.getEntity().getContent()));
            StringBuilder result = new StringBuilder();
            String line;
            while ((line = rd.readLine()) != null) result.append(line);

            System.out.println("✅ Response for ASN " + asnId + " (HTTP " + statusCode + "): " + result);

            // Check status code AND success field in JSON
            if (statusCode == 200 || statusCode == 201) {
                JSONObject jsonResponse = new JSONObject(result.toString());
                if (jsonResponse.has("success") && !jsonResponse.getBoolean("success")) {
                    return false; // mark as failed if success:false
                }
                return true; // success
            } else {
                return false; // any other status code = fail
            }

        } catch (Exception e) {
            System.out.println("❌ Exception in triggerLPNASNAPI():");
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
