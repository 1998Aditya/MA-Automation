package MA_MSG_Suite_INB;

import MA_MSG_Suite_OB.DocPathManager;
import org.apache.poi.ss.usermodel.*;
import org.openqa.selenium.*;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.FileInputStream;
import java.time.Duration;
import java.util.*;

/**
 * =========================================================
 * LPNlvlValidation
 * ---------------------------------------------------------
 * Purpose:
 * - Validate LPN-level receiving for ASN
 * - Uses ASN IDs from LPN_ASN sheet
 * =========================================================
 */
public class LPNlvlValidation {
    public static String docPathLocal ;

    private static final String DATA_EXCEL_PATH = ExcelReaderIB.DATA_EXCEL_PATH;

    public void execute(WebDriver driver, String testcase) {

        try {
            Map<String, List<String>> testcaseAsns = readAsnsByTestcase();
            List<String> asns = testcaseAsns.get(testcase);

            if (asns == null || asns.isEmpty()) {
                System.out.println("⚠ No LPN ASNs found for testcase " + testcase);
                return;
            }
            docPathLocal = IBDocPathManager.getOrCreateDocPath(ExcelReaderIB.DOC_FILEPATH, testcase);//Screenshot
            System.out.println("Path"+docPathLocal); //Screenshot Doc
            System.out.println("🔎 LPN-level validation ASNs: " + asns);

            for (String asn : asns) {
                validateLpnForAsn(driver, asn);
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void validateLpnForAsn(WebDriver driver, String asn) {

        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(20));

        try {
            WebElement searchInput = wait.until(
                    ExpectedConditions.elementToBeClickable(
                            By.xpath("//input[contains(@id,'ion-input')]"))
            );
            searchInput.clear();
            searchInput.sendKeys(asn);

            searchInput.sendKeys(Keys.ENTER);

            WebElement asnCard = wait.until(
                    ExpectedConditions.elementToBeClickable(
                            By.cssSelector("card-view[data-component-id='Card-View'] div.card-row.primary"))
            );
            asnCard.click();
            Thread.sleep(3000);
            IBDocPathManager.captureScreenshot(driver,"Created ASN"); //Screenshot
            IBDocPathManager.saveSharedDocument();                               //Screenshot Doc
            System.out.println("Created ASN screenshot done");

            WebElement relatedLinks = wait.until(
                    ExpectedConditions.elementToBeClickable(
                            By.cssSelector("button[data-component-id='relatedLinks']"))
            );
            relatedLinks.click();

            WebElement lpnBtn = wait.until(
                    ExpectedConditions.elementToBeClickable(
                            By.cssSelector("ion-item[data-component-id='LPN']"))
            );
            lpnBtn.click();
            Thread.sleep(3000);
            IBDocPathManager.captureScreenshot(driver,"Created LPNs"); //Screenshot
            IBDocPathManager.saveSharedDocument();                               //Screenshot Doc
            System.out.println("Created LPNs screenshot done");

        } catch (Exception e) {
            System.out.println("❌ LPN validation failed for ASN " + asn);
            e.printStackTrace();
        }
    }

    private Map<String, List<String>> readAsnsByTestcase() throws Exception {

        Map<String, List<String>> map = new LinkedHashMap<>();

        try (FileInputStream fis = new FileInputStream(DATA_EXCEL_PATH);
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet sheet = wb.getSheet("LPN_ASN");
            if (sheet == null)
                throw new RuntimeException("Sheet 'LPN_ASN' not found");

            Row header = sheet.getRow(0);
            int asnCol = -1;
            int tcCol = -1;

            for (int i = 0; i < header.getLastCellNum(); i++) {
                String h = header.getCell(i).getStringCellValue();
                if ("AsnId".equalsIgnoreCase(h)) asnCol = i;
                if ("Testcase".equalsIgnoreCase(h)) tcCol = i;
            }

            Iterator<Row> it = sheet.iterator();
            it.next();

            while (it.hasNext()) {
                Row r = it.next();
                String tc = getCellValue(r.getCell(tcCol));
                String asn = getCellValue(r.getCell(asnCol));
                if (!tc.matches("TST_\\d+") || asn.isEmpty()) continue;
                map.computeIfAbsent(tc, k -> new ArrayList<>()).add(asn);
            }
        }
        return map;
    }

    private String getCellValue(Cell cell) {
        return cell == null ? "" : new DataFormatter().formatCellValue(cell).trim();
    }
}
