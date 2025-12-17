package MA_MSG_Suite_OB;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;

import java.io.FileInputStream;

public class Main1_URL_Login1 {

    private WebDriver driver;
    private String environment;

    // constructor to receive driver and environment
    public Main1_URL_Login1(WebDriver driver, String environment) {
        this.driver = driver;
        this.environment = environment;
    }

    public void execute() {
        String excelPath = "C://Users//2210420//IdeaProjects//msg-runner//Login.xlsx";
        String username = "";
        String password = "";
        String BASE_URL = "";

        try (FileInputStream fis = new FileInputStream(excelPath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheet("Login");
            if (sheet == null) {
                throw new RuntimeException("Sheet 'Login' not found in Excel file!");
            }

            boolean found = false;
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                Cell envCell = row.getCell(8); // H column - Environment
                if (envCell == null) continue;

                String envValue = envCell.getStringCellValue().trim();
                if (envValue.equalsIgnoreCase(environment.trim())) {
                    BASE_URL = row.getCell(1).getStringCellValue(); // B
                    username = row.getCell(2).getStringCellValue(); // C
                    password = row.getCell(3).getStringCellValue(); // D
                    found = true;
                    System.out.println("🌍 Environment matched: " + envValue);
                    break;
                }
            }

            if (!found) {
                throw new RuntimeException("Environment '" + environment + "' not found in Excel file!");
            }

        } catch (Exception e) {
            throw new RuntimeException("Failed to read Excel: " + e.getMessage(), e);
        }

        // Navigate & login
        driver.get(BASE_URL);
        driver.manage().window().maximize();

        driver.findElement(By.name("username")).sendKeys(username);
        driver.findElement(By.name("password")).sendKeys(password);
        driver.findElement(By.name("login")).click();
    }
}
