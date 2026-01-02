package MA_MSG_Suite_INB;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.support.ui.Select;
import java.io.FileInputStream;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class URL_Login_old {

    private WebDriver driver;

    // constructor to receive driver from MainApp
    public URL_Login_old(WebDriver driver) {
        this.driver = driver;
    }

    public void execute() {
        String excelPath = "C://Users//Aditya Mishra//IdeaProjects//msg-runner//Login.xlsx";
        String username = "";
        String password = "";
        String BASE_URL = "";
        String environment = ""; // new variable

        // Read username, password, URL, and environment from Excel
        try (FileInputStream fis = new FileInputStream(excelPath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheet("Login");
            if (sheet == null) {
                throw new RuntimeException("Sheet 'Login' not found in Excel file!");
            }

            Row row = sheet.getRow(1); // second row
            BASE_URL = row.getCell(1).getStringCellValue(); // col B
            username = row.getCell(2).getStringCellValue(); // col C
            password = row.getCell(3).getStringCellValue(); // col D
            environment = row.getCell(7).getStringCellValue(); // col H (Environment)

        } catch (Exception e) {
            e.printStackTrace();
        }

        // Navigate to site
        driver.get(BASE_URL);
        driver.manage().window().maximize();

        // Perform login
        driver.findElement(By.name("username")).sendKeys(username);
        driver.findElement(By.name("password")).sendKeys(password);
        driver.findElement(By.name("login")).click();

        // Select environment (if dropdown or radio button exists)
        try {
            // Example for dropdown selection
            Select envDropdown = new Select(driver.findElement(By.id("environment")));
            envDropdown.selectByVisibleText(environment);

            // OR if it’s a clickable element:
            // driver.findElement(By.xpath("//label[text()='" + environment + "']")).click();

            System.out.println("Environment selected: " + environment);

        } catch (Exception e) {
            System.out.println("Environment selection failed: " + e.getMessage());
        }
    }
}
