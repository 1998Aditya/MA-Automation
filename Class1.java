package org.example;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.apache.poi.ss.usermodel.*;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.FileInputStream;
import java.io.IOException;
import java.time.Duration;

public class Class1 {

    // Login and return driver
    public static WebDriver login() {
        WebDriverManager.chromedriver().setup();
        WebDriver driver = new ChromeDriver();

        driver.get("https://ujdss-auth.sce.manh.com/auth/realms/maactive/protocol/openid-connect/auth?scope=openid&client_id=zuulserver.1.0.0&redirect_uri=https://ujdss.sce.manh.com/login&response_type=code");
        driver.manage().window().maximize();

        driver.findElement(By.id("username")).sendKeys("cogs");
        driver.findElement(By.id("password")).sendKeys("Cogs@123456");
        driver.findElement(By.id("kc-login")).click();

        try {
            Thread.sleep(20000); // wait for login
        } catch (InterruptedException e) {
            e.printStackTrace();
        }

        return driver;
    }

    // Read all ASNs from Excel (Sheet1, column B, skipping header in B1)
    public static String[] readASNsFromExcel() throws IOException {
        String excelPath = "C:\\Users\\2378594\\IdeaProjects\\Testcases - Copy\\Testcases - Copy\\OOdata.xlsx";
        FileInputStream fis = new FileInputStream(excelPath);
        Workbook workbook = WorkbookFactory.create(fis);

        Sheet sheet = workbook.getSheet("Sheet1");
        if (sheet == null) {
            throw new RuntimeException("Sheet 'Sheet1' not found in Excel file!");
        }

        int lastRow = sheet.getLastRowNum();
        String[] asns = new String[lastRow]; // skip header row

        for (int i = 1; i <= lastRow; i++) { // start from row 1 (Excel row 2)
            Row row = sheet.getRow(i);
            if (row != null) {
                Cell cell = row.getCell(1); // column B = index 1
                if (cell != null && cell.getCellType() == CellType.STRING) {
                    asns[i - 1] = cell.getStringCellValue();
                }
            }
        }

        workbook.close();
        fis.close();
        return asns;
    }

    static void SearchandOpenWMMobie(WebDriver driver) throws IOException {
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(15));
        JavascriptExecutor js = (JavascriptExecutor) driver;

        try {
            // Open menu
            WebElement shadowHost = driver.findElement(By.cssSelector("ion-button.menu-toggle-button"));
            SearchContext shadowRoot = (SearchContext) js.executeScript("return arguments[0].shadowRoot", shadowHost);
            shadowRoot.findElement(By.cssSelector("button.button-native")).click();

            WebElement searchInput = wait.until(ExpectedConditions.elementToBeClickable(
                    By.xpath("//input[@placeholder='Search Menu...']")));
            searchInput.clear();
            searchInput.sendKeys("ASNS");

            WebElement asnsButton = wait.until(ExpectedConditions.elementToBeClickable(By.id("ASN")));
            asnsButton.click();

            Thread.sleep(5000);

            // Expand filter section
            WebElement filterBtnHost = wait.until(ExpectedConditions.elementToBeClickable(
                    By.xpath("(//ion-button[contains(@class,'toggle-button')])[3]")));
            js.executeScript("arguments[0].shadowRoot.querySelector('.button-inner').click();", filterBtnHost);

            Thread.sleep(3000);

            // Dropdown expand
            WebElement dropdownHost = wait.until(ExpectedConditions.presenceOfElementLocated(
                    By.cssSelector("ion-button[data-component-id='ASN-ASN-chevron-down']")
            ));
            WebElement dropdownButton = (WebElement) js.executeScript(
                    "return arguments[0].shadowRoot.querySelector('button.button-native')",
                    dropdownHost
            );
            js.executeScript("arguments[0].scrollIntoView(true);", dropdownButton);
            wait.until(ExpectedConditions.elementToBeClickable(dropdownButton)).click();
            System.out.println("ASN chevron-down dropdown clicked successfully.");

            Thread.sleep(3000);

            // Click input field first
            WebElement orderPlanningRunInputField = wait.until(
                    ExpectedConditions.elementToBeClickable(By.id("ion-input-8"))
            );
            orderPlanningRunInputField.click();
            System.out.println("Input field clicked successfully.");

            // Read all ASNs from Excel
            String[] asns = readASNsFromExcel();

            // Enter each ASN one by one
            for (String asnValue : asns) {
                if (asnValue != null && !asnValue.isEmpty()) {
                    // Clear field before entering
                    orderPlanningRunInputField.clear();

                    // Use JS to set value and trigger Angular/Ionic binding
                    js.executeScript(
                            "arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('input'));",
                            orderPlanningRunInputField, asnValue
                    );

                    System.out.println("Entered ASN: " + asnValue);

                    // Press ENTER to confirm
                    orderPlanningRunInputField.sendKeys(Keys.ENTER);

                    Thread.sleep(2000); // wait between entries

                    // ✅ Select ASN card
                    WebElement asnCard = wait.until(ExpectedConditions.elementToBeClickable(
                            By.cssSelector("card-view[data-component-id='Card-View'] div.card-row.primary")
                    ));
                    asnCard.click();
                    System.out.println("ASN card selected successfully.");

                    Thread.sleep(2000);

                    // ✅ Click Related Links button
                    WebElement relatedLinksBtn = wait.until(ExpectedConditions.elementToBeClickable(
                            By.cssSelector("button[data-component-id='relatedLinks']")
                    ));
                    relatedLinksBtn.click();
                    System.out.println("Related Links button clicked successfully.");

                    Thread.sleep(2000);

                    // ✅ Click LPN button
                    WebElement lpnButton = wait.until(ExpectedConditions.elementToBeClickable(
                            By.cssSelector("ion-item[data-component-id='LPN']")
                    ));
                    lpnButton.click();
                    System.out.println("LPN button clicked successfully.");

                    Thread.sleep(2000);
                }
            }

        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    public static void main(String[] args) throws IOException {
        WebDriver driver = login();
        SearchandOpenWMMobie(driver);
    }

    public static void execute() {
    }
}
