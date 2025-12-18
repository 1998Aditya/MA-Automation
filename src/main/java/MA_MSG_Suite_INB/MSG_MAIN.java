/*package MA_MSG_Suite_INB;

//import MA_MSG_Suite_INB.URL_Login;
import io.github.bonigarcia.wdm.WebDriverManager;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class MSG_MAIN {
    public static void main(String[] args) throws InterruptedException {
        System.out.println("Starting the application sequence...");


        // 3. Step 3: Create ASN (API based, no Selenium)
        MSG_Item_ASN_Creation step3 = new MSG_Item_ASN_Creation();
        step3.execute();

        // 4. Step 4: RCV Item lvl ASN (API based, no Selenium)
        Thread.sleep(10000);
        MSG_ReportRCV_itemlvl step4 = new MSG_ReportRCV_itemlvl();
        step4.execute();

        // 5. Step 5: RCV LPN lvl ASN (API based, no Selenium)
        Thread.sleep(10000);
        MSG_ReportRCV_LPNLvl step5 = new MSG_ReportRCV_LPNLvl();
        step5.execute();

        // 6. Step 6:GetConditionCode- IF CR or FI then it will trigger BOXDelivered msg
        Thread.sleep(5000);
        GetConditionCode step6 = new GetConditionCode();
        step6.execute();

        // 7. Step 7: BoxDelivered
        Thread.sleep(5000);
        BoxDelivered step7 = new BoxDelivered();
        step7.execute();


        // 8. Step 8: RemoveConditionCode
        Thread.sleep(5000);
        RemoveConditionCode step8 = new RemoveConditionCode();
        step8.execute();

 // for induct iLPN run step 1,2,9
        // 1. Setup ChromeDriver (replace path if needed)
        WebDriverManager.chromedriver().setup();
        WebDriver driver = new ChromeDriver();
        driver.manage().window().maximize();

        // 2. Step 1: Login
        URL_Login step1 = new URL_Login(driver);
        step1.execute();

        // 9. Step 9: InductiLPN
        Thread.sleep(5000);
        InductiLPN_MFS step9 = new InductiLPN_MFS(driver);
        step9.execute();

    // 10. Step 10: iLPNToted
        Thread.sleep(5000);
        iLPNToted step10 = new iLPNToted();
        step10.execute();


        Thread.sleep(200);
        System.out.println("Application sequence finished.");
        //driver.quit();
    }
}
*/