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
import org.json.JSONTokener;

import java.io.FileInputStream;
import java.io.InputStream;
import java.io.File;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;

/**
 * LocationBarcodeService
 *
 * Helper to call:
 * POST - {{url}}dcinventory/api/dcinventory/location/quickSearch
 *
 * Example payloads:
 * If LocationId:
 * {
 *   "Query": "LocationId in ('LOC123')",
 *   "Size": 100
 * }
 *
 * Or by TaskMovementZone:
 * {
 *   "Query": "TaskMovementZoneId in ('MHE_TASKMOVE_UNAVAIL')",
 *   "Size": 100
 * }
 *
 * Returns the first LocationBarcode found (or empty string if none).
 *
 * Keep existing comments — do not remove them.
 */
public class LocationBarcodeService {

    // Login file path (from your central reader)
    private static final String LOGIN_EXCEL_PATH = ExcelReaderIB.LOGIN_EXCEL_PATH;

    private static String LOGIN_URL;
    private static String BASE_URL;
    private static String USERNAME;
    private static String PASSWORD;
    private static String CLIENT_ID;
    private static String CLIENT_SECRET;
    private static String SELECTED_ORG;
    private static String SELECTED_LOC;

    private static String QUICKSEARCH_URL;

    static {
        loadLoginConfig();
    }

    // ✅ Load login info from Login.xlsx
    private static void loadLoginConfig() {
        try (FileInputStream fis = new FileInputStream(new File(LOGIN_EXCEL_PATH));
             Workbook workbook = WorkbookFactory.create(fis)) {

            Sheet sheet = workbook.getSheet("Login");
            if (sheet == null) {
                System.out.println("❌ Login sheet missing in " + LOGIN_EXCEL_PATH);
                return;
            }
            Row row = sheet.getRow(1);
            if (row == null) {
                System.out.println("❌ Login row (1) missing in " + LOGIN_EXCEL_PATH);
                return;
            }

            LOGIN_URL = getCellString(row.getCell(0));
            BASE_URL  = getCellString(row.getCell(1));
            USERNAME  = getCellString(row.getCell(2));
            PASSWORD  = getCellString(row.getCell(3));
            CLIENT_ID = getCellString(row.getCell(4));
            CLIENT_SECRET = getCellString(row.getCell(5));
            SELECTED_ORG = getCellString(row.getCell(6));
            SELECTED_LOC = getCellString(row.getCell(7));

            if (BASE_URL == null) BASE_URL = "";
            if (!BASE_URL.endsWith("/")) BASE_URL += "/";

            QUICKSEARCH_URL = BASE_URL + "dcinventory/api/dcinventory/location/quickSearch";

            System.out.println("Loaded Login config for LocationBarcodeService:");
            System.out.println("  LOGIN_URL     = " + LOGIN_URL);
            System.out.println("  QUICKSEARCH_URL = " + QUICKSEARCH_URL);

        } catch (Exception e) {
            System.out.println("❌ Error loading Login.xlsx in LocationBarcodeService");
            e.printStackTrace();
        }
    }

    // ✅ Get authentication token (same pattern as other services)
    private static String getToken() {
        try {
            HttpClient client = HttpClients.createDefault();
            HttpPost post = new HttpPost(LOGIN_URL);

            // Use url-encoded form entity for token request
            List<NameValuePair> params = new ArrayList<>();
            params.add(new BasicNameValuePair("grant_type", "password"));
            params.add(new BasicNameValuePair("username", USERNAME));
            params.add(new BasicNameValuePair("password", PASSWORD));
            params.add(new BasicNameValuePair("client_id", CLIENT_ID));
            params.add(new BasicNameValuePair("client_secret", CLIENT_SECRET));

            post.setEntity(new UrlEncodedFormEntity(params, StandardCharsets.UTF_8));
            post.setHeader("Content-Type", "application/x-www-form-urlencoded");

            ClassicHttpResponse resp = (ClassicHttpResponse) client.execute(post);
            InputStream is = resp.getEntity() != null ? resp.getEntity().getContent() : null;
            if (is == null) {
                System.out.println("❌ Empty token response stream");
                return "";
            }
            String result = new String(is.readAllBytes(), StandardCharsets.UTF_8);

            JSONObject json = new JSONObject(result);
            String token = json.optString("access_token", "");
            if (token.isEmpty()) {
                System.out.println("⚠ Token response did not contain access_token. Full response:\n" + result);
            }
            return token;

        } catch (Exception e) {
            System.out.println("❌ Token generation failed in LocationBarcodeService");
            e.printStackTrace();
            return "";
        }
    }

    // ======================================================
    //           Public convenience methods
    // ======================================================

    /**
     * Query quickSearch by LocationId and return the first LocationBarcode found (or empty string).
     */
    public static String getLocationBarcodeByLocationId(String locationId) {
        if (locationId == null || locationId.trim().isEmpty()) return "";
        // ensure single-quoted value(s) inside parentheses
        String val = locationId.trim();
        if (!val.startsWith("'")) val = "'" + val;
        if (!val.endsWith("'")) val = val + "'";
        String query = "LocationId in (" + val + ")";
        return runQuickSearchAndExtractBarcode(query);
    }

    /**
     * Query quickSearch by TaskMovementZoneId and return the first LocationBarcode found (or empty string).
     */
    public static String getLocationBarcodeByTaskMovementZone(String zoneId) {
        if (zoneId == null || zoneId.trim().isEmpty()) return "";
        String val = zoneId.trim();
        if (!val.startsWith("'")) val = "'" + val;
        if (!val.endsWith("'")) val = val + "'";
        String query = "TaskMovementZoneId in (" + val + ")";
        return runQuickSearchAndExtractBarcode(query);
    }

    // ======================================================
    //           Internal call + parsing logic
    // ======================================================

    private static String runQuickSearchAndExtractBarcode(String query) {
        try {
            String token = getToken();
            if (token == null || token.isEmpty()) {
                System.out.println("❌ No token received in LocationBarcodeService");
                return "";
            }

            JSONObject payload = new JSONObject();
            payload.put("Query", query);
            payload.put("Size", 100);

            HttpClient client = HttpClients.createDefault();
            HttpPost post = new HttpPost(QUICKSEARCH_URL);

            post.setHeader("Authorization", "Bearer " + token);
            post.setHeader("SelectedOrganization", SELECTED_ORG);
            post.setHeader("SelectedLocation", SELECTED_LOC);
            post.setHeader("Content-Type", "application/json");

            post.setEntity(new StringEntity(payload.toString(), StandardCharsets.UTF_8));

            ClassicHttpResponse resp = (ClassicHttpResponse) client.execute(post);
            InputStream is = resp.getEntity() != null ? resp.getEntity().getContent() : null;
            if (is == null) {
                System.out.println("⚠ Empty response from quickSearch for query: " + query);
                return "";
            }
            String result = new String(is.readAllBytes(), StandardCharsets.UTF_8);

            // parse robustly (JSONObject or JSONArray)
            Object parsed;
            try {
                parsed = new JSONTokener(result).nextValue();
            } catch (Exception pe) {
                System.out.println("❌ Failed to parse JSON quickSearch response: " + pe.getMessage());
                System.out.println("Full response:\n" + result);
                return "";
            }

            JSONArray arr = null;
            if (parsed instanceof JSONArray) {
                arr = (JSONArray) parsed;
            } else if (parsed instanceof JSONObject) {
                JSONObject j = (JSONObject) parsed;
                if (j.has("items") && j.get("items") instanceof JSONArray) {
                    arr = j.getJSONArray("items");
                } else if (j.has("data") && j.get("data") instanceof JSONArray) {
                    arr = j.getJSONArray("data");
                } else if (j.has("result") && j.get("result") instanceof JSONArray) {
                    arr = j.getJSONArray("result");
                } else {
                    // fallback: first JSONArray value in object
                    for (String k : j.keySet()) {
                        Object o = j.get(k);
                        if (o instanceof JSONArray) {
                            arr = (JSONArray) o;
                            break;
                        }
                    }
                }
            }

            if (arr == null || arr.isEmpty()) {
                System.out.println("⚠ No locations found for query: " + query);
                System.out.println("Full response:\n" + result);
                return "";
            }

            JSONObject loc = arr.getJSONObject(0);

            // common field LocationBarcode (case-sensitive)
            if (loc.has("LocationBarcode") && !loc.optString("LocationBarcode", "").isEmpty()) {
                return loc.getString("LocationBarcode").trim();
            }
            // fallback keys (case-insensitive)
            for (String key : loc.keySet()) {
                if ("locationbarcode".equalsIgnoreCase(key) && !loc.optString(key, "").isEmpty()) {
                    return loc.optString(key, "").trim();
                }
            }

            // maybe nested under "Extended" or similar
            if (loc.has("Extended") && loc.get("Extended") instanceof JSONObject) {
                JSONObject ext = loc.getJSONObject("Extended");
                for (String key : ext.keySet()) {
                    if ("locationbarcode".equalsIgnoreCase(key) && !ext.optString(key, "").isEmpty()) {
                        return ext.optString(key, "").trim();
                    }
                }
            }

            System.out.println("⚠ Location found but no LocationBarcode field present. Location object:\n" + loc.toString(4));
            return "";

        } catch (Exception e) {
            System.out.println("❌ Error in runQuickSearchAndExtractBarcode()");
            e.printStackTrace();
            return "";
        }
    }

    // small helper to safely extract string from Excel cell (prevents NPE)
    private static String getCellString(Cell cell) {
        if (cell == null) return "";
        DataFormatter fmt = new DataFormatter();
        return fmt.formatCellValue(cell).trim();
    }
}
