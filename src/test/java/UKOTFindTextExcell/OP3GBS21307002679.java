package UKOTFindTextExcell;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class OP3GBS21307002679 {


    static XSSFWorkbook workbook;
    static XSSFSheet sheet;
    static String URL;
    static String TITLE;
    static WebDriver driver ;
    static String TEXT;
    static FileOutputStream fos;
    static String projectPath = System.getProperty("user.dir");
    static FileInputStream fis;
    static String XLS = "/exel/OP3GBS21307002679/data.xlsx";
    static String FILE = projectPath + XLS;
    static Row row;
    static WebElement okThanks;
    static WebElement removeButton;
    static JavascriptExecutor js = (JavascriptExecutor)driver;
    static String theSelector;

    static int TESTING_ROW = 0;
    static int TESTING_URL_CELL = 1;
    static int TESTING_DATA_CEL = 2;
    static int TEST_RESULT_CEL =  2;
    static int TEST_RESULT_CEL3 =  3;
    static int TEST_RESULT_CEL4 =  4;
    static int TEST_RESULT_CEL5 =  5;
    static int TEST_RESULT_CEL6 =  6;
    static int TEST_RESULT_CEL7 =  7;
    static int TEST_RESULT_CEL8 =  8;
    static int TEST_RESULT_CEL9 =  9;
    static int TEST_RESULT_CEL10 =  10;


    static int n;
    static int drop;

//    static WebElement okThanks = driver.findElement(By.cssSelector("#header > div.seuk-popup-cart--wrap > div > div > div.seuk-popup-cart--info > span"));
//    static WebElement removeButton = driver.findElement(By.cssSelector("#app > div.site-wrapper.siteHeader.scrollHeight > main > div > div > div.container.checkout-cart-page-holder.page-holder > div.grid > div.shopping-cart-holder > div.shopping-cart-items-holder.cart-items > div:nth-child(2) > div > div.sc-product-meta-details > div.sc-product-details > div.sc-product-meta-data > p > span.for-desktop.data-omni-remove > button\n"));

//    static int TESTING_ROW;
//    static int TESTING_URL_CELL;
//    static int TESTING_DATA_CEL;
//    static int TEST_RESULT_CEL;
//    static int TESTING_ITEM_URL_CELL;
//    static int n;

    public static void main(String[] args) throws Exception {

        n = 1;

        for (int i = 0; i <10; i++) {
            activityTest();
        }
        System.out.println("all done now!");
        quit();

    }


    public static void activityTest() throws Exception {
        openBrowser();

        goToPage(TESTING_ROW + n, TESTING_DATA_CEL);
        Thread.sleep(4000);
        acceptLogin();

        goToPage(TESTING_ROW + n, TEST_RESULT_CEL8);
        Thread.sleep(4000);
        goToPage(TESTING_ROW + n, TESTING_URL_CELL);
        Thread.sleep(6000);


        findExpectedText(TESTING_ROW + n, TEST_RESULT_CEL3);
        Thread.sleep(1000);
        System.out.println(n);
        findExpectedText(TESTING_ROW + n, TEST_RESULT_CEL4);
        System.out.println(n);
        findExpectedTextAndLink(TESTING_ROW + n, TEST_RESULT_CEL4);
        System.out.println(n);


        quit();
        n = n + 1;


    }

    public static void acceptLogin () throws Exception{
        driver.findElement(By.cssSelector("#truste-consent-button")).click();
        Thread.sleep(4000);
        driver.findElement(By.cssSelector("body > app-root > cx-storefront > main > cx-page-layout > cx-page-slot.MultistoreLogin.has-components > app-sign-in > div > div.sign-in-component__content-footer > button")).click();
        Thread.sleep(3000);

        driver.findElement(By.cssSelector("#iptLgnPlnID")).sendKeys("mykola.r@cheil.com");
        Thread.sleep(2000);
        driver.findElement(By.cssSelector("#iptLgnPlnPD")).sendKeys("M88kolasas!");
        driver.findElement(By.cssSelector("#iptLgnPlnPD")).sendKeys(Keys.ENTER);
        Thread.sleep(2000);
        driver.findElement(By.cssSelector("body > div.wrapper.ng-scope > main > div > div.container > div.content > div > label > span")).click();
        Thread.sleep(2000);
        driver.findElement(By.cssSelector("#btnNotNow")).click();
        Thread.sleep(3000);
//        Thread.sleep(1000);
//        driver.findElement(By.cssSelector("#header-main > div.main-nav-left > div.main-nav-logo > a > mat-icon > svg")).click();
    }
    public static void scrollClick () throws Exception {

        JavascriptExecutor js = (JavascriptExecutor) driver;
        Thread.sleep(2000);
        js.executeScript("document.body.style.zoom = '0.5'");
//        js.executeScript("window.scrollTo (0, " + drop + ") ");
    }
    public static void openBrowser() throws Exception {
        driver = new ChromeDriver();
        driver.manage().window().maximize();
    }
    public static void getTheSelector( int rowNum, int cellNum) throws Exception {
        try {
            workbook = new XSSFWorkbook(FILE);
            sheet = workbook.getSheetAt(0);
            theSelector = sheet.getRow(rowNum).getCell(cellNum).getStringCellValue();
        } catch (Exception exp) {
            System.out.println(exp.getMessage());
            System.out.println(exp.getCause());
            exp.printStackTrace();
        }

        System.out.println("theSelector found");}
    public static void goToPage( int rowNum, int cellNum) throws Exception {
        try {
            workbook = new XSSFWorkbook(FILE);
            sheet = workbook.getSheetAt(0);
            URL = sheet.getRow(rowNum).getCell(cellNum).getStringCellValue();
        } catch (Exception exp) {
            System.out.println(exp.getMessage());
            System.out.println(exp.getCause());
            exp.printStackTrace();
        }
        driver.get(URL);
        System.out.println("URL found - "+URL);}
    public static void getN(int rowNum, int cellNum) {
        try {
            sheet = workbook.getSheetAt(0);
            n = (int) sheet.getRow(rowNum).getCell(cellNum).getNumericCellValue();
        } catch (Exception exp) {
            System.out.println(exp.getMessage());
            System.out.println(exp.getCause());
            exp.printStackTrace();
        }
    }

    public static void writeFile(XSSFWorkbook workbook,String FILE ) throws IOException, InterruptedException {
        FileOutputStream fos = new FileOutputStream(FILE);
        workbook.write(fos);
        fos.close();
        workbook.close();
        System.out.println("File Writen");
    }


    public static void findExpectedTextAndLink(int rowNum, int colNum) throws IOException, InterruptedException{

       try{
           getTheSelector(TESTING_ROW + n, colNum);
           System.out.println(theSelector);

           WebElement meta = driver.findElement(By.cssSelector(theSelector));

           String meta_text = meta.getAttribute("href");
           meta.click();
           Thread.sleep(4000);

           URL = driver.getCurrentUrl();
//           TITLE = driver.getTitle();

           System.out.println("Founud -  "+meta_text);
           System.out.println("Founud -  "+URL);
//           System.out.println("Founud -  "+TITLE);
           fis = new FileInputStream(FILE);
           workbook = new XSSFWorkbook(fis);
           sheet = workbook.getSheetAt(0);
           row = sheet.getRow(rowNum);
           row.createCell(colNum+2).setCellValue(meta_text);
           row.createCell(colNum + 3).setCellValue(URL);
//           row.createCell(colNum + 3).setCellValue(TITLE);

       } catch (Exception exp) {
        System.out.println(exp.getMessage());
        System.out.println(exp.getCause());
        exp.printStackTrace();
           fis = new FileInputStream(FILE);
           workbook = new XSSFWorkbook(fis);
           sheet = workbook.getSheetAt(0);
           row = sheet.getRow(rowNum);
           row.createCell(colNum+2).setCellValue("FAIL");
    }
        writeFile(workbook, FILE);



    }
    public static void findExpectedText(int rowNum, int colNum) throws IOException, InterruptedException{

        try{
            getTheSelector(TESTING_ROW + n, colNum);
            System.out.println(theSelector);

            WebElement meta = driver.findElement(By.cssSelector(theSelector));

            String meta_text = meta.getText();

            System.out.println("Founud -  "+meta_text);
            fis = new FileInputStream(FILE);
            workbook = new XSSFWorkbook(fis);
            sheet = workbook.getSheetAt(0);
            row = sheet.getRow(rowNum);
            row.createCell(colNum+2).setCellValue(meta_text);

        } catch (Exception exp) {
            System.out.println(exp.getMessage());
            System.out.println(exp.getCause());
            exp.printStackTrace();
            fis = new FileInputStream(FILE);
            workbook = new XSSFWorkbook(fis);
            sheet = workbook.getSheetAt(0);
            row = sheet.getRow(rowNum);
            row.createCell(colNum+2).setCellValue("FAIL");
        }
        writeFile(workbook, FILE);
    }
    public static void quit(){
        driver.quit();
    }



}
