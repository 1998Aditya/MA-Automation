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
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;

/**
 * ItemBarcodeService
 *
 * Fetches item master via API and returns PrimaryBarCode or MAUJDSDefaultEANBarcode (fallback).
 *
 * - Login file path (from your central reader)
 * - Uses HttpClient 5.x
 *
 * Keep existing comments — do not remove them.
 */
public class ItemBarcodeService {

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

    private static String SEARCH_URL;

    static {
        loadLoginConfig();
    }

    // ✅ Load login info from Login.xlsx
    private static void loadLoginConfig() {
        try (FileInputStream fis = new FileInputStream(LOGIN_EXCEL_PATH);
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

            SEARCH_URL = BASE_URL + "item-master/api/item-master/item/search";

            System.out.println("Loaded Login config for ItemBarcodeService:");
            System.out.println("  LOGIN_URL = " + LOGIN_URL);
            System.out.println("  SEARCH_URL = " + SEARCH_URL);

        } catch (Exception e) {
            System.out.println("❌ Error loading Login.xlsx in ItemBarcodeService");
            e.printStackTrace();
        }
    }

    // ✅ Get authentication token
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
            System.out.println("❌ Token generation failed in ItemBarcodeService");
            e.printStackTrace();
            return "";
        }
    }

    // ======================================================
    //               ★ MAIN PUBLIC FUNCTION ★
    // ======================================================
    public static String getItemBarcode(String itemId) {

        try {
            String token = getToken();
            if (token == null || token.isEmpty()) {
                System.out.println("❌ No token received in ItemBarcodeService");
                return "";
            }

            // Build payload exactly as required by API
            JSONObject payload = new JSONObject();
            // Ensure itemId is quoted if necessary (API expects values in parentheses possibly comma-separated)
            String qItem = itemId;
            if (!qItem.startsWith("'") && !qItem.endsWith("'")) {
                qItem = "'" + qItem + "'";
            }
            payload.put("Query", "ItemId IN (" + qItem + ")");

            HttpClient client = HttpClients.createDefault();
            HttpPost post = new HttpPost(SEARCH_URL);

            post.setHeader("Authorization", "Bearer " + token);
            post.setHeader("SelectedOrganization", SELECTED_ORG);
            post.setHeader("SelectedLocation", SELECTED_LOC);
            post.setHeader("Content-Type", "application/json");

            post.setEntity(new StringEntity(payload.toString(), StandardCharsets.UTF_8));

            ClassicHttpResponse resp = (ClassicHttpResponse) client.execute(post);
            InputStream is = resp.getEntity() != null ? resp.getEntity().getContent() : null;
            if (is == null) {
                System.out.println("⚠ Empty response when searching item: " + itemId);
                return "";
            }
            String result = new String(is.readAllBytes(), StandardCharsets.UTF_8);

            // Use JSONTokener to parse either JSONObject or JSONArray safely
            Object parsed = null;
            try {
                parsed = new JSONTokener(result).nextValue();
            } catch (Exception pe) {
                System.out.println("❌ Failed to parse JSON response: " + pe.getMessage());
                System.out.println("Full response:\n" + result);
                return "";
            }

            JSONArray items = null;

            if (parsed instanceof JSONArray) {
                items = (JSONArray) parsed;
            } else if (parsed instanceof JSONObject) {
                JSONObject json = (JSONObject) parsed;
                // The endpoint may return the array under different keys. Try several fallbacks.
                if (json.has("items") && json.get("items") instanceof JSONArray) {
                    items = json.getJSONArray("items");
                } else if (json.has("data") && json.get("data") instanceof JSONArray) {
                    items = json.getJSONArray("data");
                } else if (json.has("result") && json.get("result") instanceof JSONArray) {
                    items = json.getJSONArray("result");
                } else {
                    // Try to find the first JSONArray anywhere (best-effort)
                    for (String key : json.keySet()) {
                        Object o = json.get(key);
                        if (o instanceof JSONArray) {
                            items = (JSONArray) o;
                            break;
                        }
                    }
                }
            } else {
                System.out.println("⚠ Unexpected JSON root type: " + (parsed == null ? "null" : parsed.getClass().getName()));
                System.out.println("Full response:\n" + result);
                return "";
            }

            if (items == null || items.isEmpty()) {
                System.out.println("⚠ No item found for ID: " + itemId + ". Full response:\n" + result);
                return "";
            }

            JSONObject item = items.getJSONObject(0);

            // Try PrimaryBarCode first (case-sensitive common field)
            if (item.has("PrimaryBarCode") && !item.optString("PrimaryBarCode", "").isEmpty()) {
                return item.getString("PrimaryBarCode").trim();
            }

            // Try some case-insensitive/variant fallbacks
            for (String key : item.keySet()) {
                if ("primarybarcode".equalsIgnoreCase(key) && !item.optString(key, "").isEmpty()) {
                    return item.optString(key, "").trim();
                }
            }

            // Then extended barcode fallback (may be nested under "Extended" or "extended")
            if (item.has("Extended") && item.get("Extended") instanceof JSONObject) {
                JSONObject ext = item.getJSONObject("Extended");
                if (ext.has("MAUJDSDefaultEANBarcode") && !ext.optString("MAUJDSDefaultEANBarcode", "").isEmpty()) {
                    return ext.getString("MAUJDSDefaultEANBarcode").trim();
                }
                // case-insensitive fallback
                for (String key : ext.keySet()) {
                    if ("maujdsdefaulteanbarcode".equalsIgnoreCase(key) && !ext.optString(key, "").isEmpty()) {
                        return ext.optString(key, "").trim();
                    }
                }
            }
            if (item.has("extended") && item.get("extended") instanceof JSONObject) {
                JSONObject ext = item.getJSONObject("extended");
                if (ext.has("MAUJDSDefaultEANBarcode") && !ext.optString("MAUJDSDefaultEANBarcode", "").isEmpty()) {
                    return ext.getString("MAUJDSDefaultEANBarcode").trim();
                }
                for (String key : ext.keySet()) {
                    if ("maujdsdefaulteanbarcode".equalsIgnoreCase(key) && !ext.optString(key, "").isEmpty()) {
                        return ext.optString(key, "").trim();
                    }
                }
            }

            // If we got here, we couldn't find a barcode
            System.out.println("⚠ Item found but no barcode fields present for item: " + itemId + ". Item object:\n" + item.toString(4));
            return "";

        } catch (Exception e) {
            System.out.println("❌ Error in getItemBarcode()");
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
