import io.github.bonigarcia.wdm.WebDriverManager;
import lombok.SneakyThrows;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.edge.EdgeDriver;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import java.util.concurrent.TimeUnit;

public class DemoTest {
    WebDriver driver;
//    Excel excel;
    String url = "https://opensource-demo.orangehrmlive.com/web/index.php/auth/login";
    Map<String, Object[]> testNGResults;

    XSSFWorkbook workbook;
    XSSFSheet sheet;
    public void login(String username, String password)  {
        driver.findElement(By.name("username")).sendKeys(username);
        driver.findElement(By.name("password")).sendKeys(password);
        driver.findElement(By.xpath("//button[@type='submit']")).click();
    }

    @SneakyThrows
    @Test(priority = 1)
    public void testLoginWithIncorrectUserNameAndPassword(){
        driver.get(url);
        login("Admin", "Password");
        if (driver.getCurrentUrl().equals("https://opensource-demo.orangehrmlive.com/web/index.php/auth/login")) {
            WebElement message = driver.findElement(By.cssSelector(".oxd-alert-content-text"));
            if (message.isDisplayed()){
                if(message.getText().equals("Invalid credentials"))
                    testNGResults.put("2", new Object[]{"1", "Admin", "Password", "Hiển thị thông tin đăng nhập không chính xác và đăng nhập không thành công", "Passed"});
            } else testNGResults.put("2", new Object[]{"1", "Admin", "Password", "Hiển thị thông tin đăng nhập không chính xác và đăng nhập không thành công", "Failed"});
        } else testNGResults.put("2", new Object[]{"1", "Admin", "Password", "Hiển thị thông tin đăng nhập không chính xác và đăng nhập không thành công", "Failed"});
    }

    @SneakyThrows
    @Test(priority = 2)
    public void testLoginWithNullUserNameAndPassword(){
        driver.get(url);
        login("", "");
        if (driver.getCurrentUrl().equals("https://opensource-demo.orangehrmlive.com/web/index.php/auth/login")) {
            WebElement message = driver.findElement(By.cssSelector(".oxd-input-group__message"));
            if (message.isDisplayed()){
                if(message.getText().equals("Required"))
                    testNGResults.put("3", new Object[]{"2", "", "", "Hiển thị thông báo Required và đăng nhập không thành công", "Passed"});
            } else testNGResults.put("3", new Object[]{"2", "", "", "Hiển thị thông báo Required và đăng nhập không thành công", "Failed"});
        } else testNGResults.put("3", new Object[]{"2", "", "", "Hiển thị thông báo Required và đăng nhập không thành công", "Failed"});
    }

    @SneakyThrows
    @Test(priority = 3)
    public void testLoginWithNullUserName(){
        driver.get(url);
        login("", "Password");
        if (driver.getCurrentUrl().equals("https://opensource-demo.orangehrmlive.com/web/index.php/auth/login")) {
            WebElement message = driver.findElement(By.cssSelector(".oxd-input-group__message"));
            if (message.isDisplayed()){
                if(message.getText().equals("Required"))
                    testNGResults.put("4", new Object[]{"3", "", "Password", "Hiển thị thông báo Required và đăng nhập không thành công", "Passed"});
            } else testNGResults.put("4", new Object[]{"3", "", "Password", "Hiển thị thông báo Required và đăng nhập không thành công", "Failed"});
        } else testNGResults.put("4", new Object[]{"3", "", "Password", "Hiển thị thông báo Required và đăng nhập không thành công", "Failed"});
    }

    @SneakyThrows
    @Test(priority = 4)
    public void testLoginWithNullPassword(){
        driver.get(url);
        login("Admin", "");
        if (driver.getCurrentUrl().equals("https://opensource-demo.orangehrmlive.com/web/index.php/auth/login")) {
            WebElement message = driver.findElement(By.cssSelector(".oxd-input-group__message"));
            if (message.isDisplayed()){
                if(message.getText().equals("Required"))
                    testNGResults.put("5", new Object[]{"4", "Admin", "", "Hiển thị thông báo Required và đăng nhập không thành công", "Passed"});
            } else testNGResults.put("5", new Object[]{"4", "Admin", "", "Hiển thị thông báo Required và đăng nhập không thành công", "Failed"});
        } else testNGResults.put("5", new Object[]{"4", "Admin", "", "Hiển thị thông báo Required và đăng nhập không thành công", "Failed"});
    }

    @Test(priority = 5)
    @SneakyThrows
    public void testLoginWithCorrectUserNameAndPassword(){
        driver.get(url);
        login("Admin", "admin123");
        if (driver.getCurrentUrl().equals("https://opensource-demo.orangehrmlive.com/web/index.php/dashboard/index"))
            testNGResults.put("6", new Object[]{"5", "Admin", "admin123", "Đăng nhập thành công", "Passed"});
        else testNGResults.put("6", new Object[]{"5", "Admin", "admin123", "Đăng nhập thành công", "Failed"});
        Thread.sleep(1000);
    }

    @BeforeClass
    public void beforeClass() throws Exception {
        //cài đặt driver
        WebDriverManager.edgedriver().setup();

        //cài đặt TestNG
        testNGResults = new LinkedHashMap<String, Object[]>();
        testNGResults.put("1", new Object[]{"Test Step", "Username", "Password", "Expected Output", "Result"});

        //Khai báo file excel result  testcases
        workbook = new XSSFWorkbook();
        sheet = workbook.createSheet("TestcaseLogin");

        try{
            driver = new EdgeDriver();
            driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
        } catch (Exception e) {
            System.out.println("Cannot run web driver");
        }
    }

    @AfterClass    public void afterClass(){
        Set<String> keyset = testNGResults.keySet();
        int rownum = 0;
        for (String key : keyset){
            Row row = sheet.createRow(rownum++);
            Object[] objects = testNGResults.get(key);
            int cellnum=0;
            for (Object obj : objects){
                Cell cell = row.createCell(cellnum++);
                if (obj instanceof String){
                    cell.setCellValue((String)obj);
                } else if ( obj instanceof Integer){
                    cell.setCellValue((Integer)obj);
                }

                System.out.println((String) obj);
            }
        }
        try{
            FileOutputStream out = new FileOutputStream(new File("src/test/resources/Results.xlsx"));
            workbook.write(out);
            out.close();
            System.out.println("Save test results to excel files succesfully!");
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        driver.close();
    }
}
