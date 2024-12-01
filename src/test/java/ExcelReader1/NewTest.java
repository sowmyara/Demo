package ExcelReader1;


import java.io.IOException;
import java.util.concurrent.TimeUnit;


import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class NewTest {
	
	WebDriver d;
	
	
  @BeforeClass
  public void f() {
	  
	  System.setProperty("webdriver.chrome.drive", "D:\\Selenium Webdriver\\Chrome driver.exe");
		
		WebDriver d = new ChromeDriver();
		 d.get("https://demoqa.com/automation-practice-form");
		 d.manage().timeouts().implicitlyWait(3, TimeUnit.SECONDS);
		d.manage().window().maximize();
	//double email, int mobile, String subjects, String c_add, String result
  }
  @Test(dataProvider="formFields")
  public void fields(String Firstname,String ln)
  {
	  
	 
	  
	  d.findElement(By.xpath("//input[@id=\"firstName\"]")).sendKeys(Firstname);
  }
  
  
  
  @DataProvider(name ="formFields")
  public Object[][] getData() throws IOException   
  {
	  
	  String filepath = "D:\\\\POCDemo.xlsx";
	  
	  XLUtility xlutil = new  XLUtility(filepath);
		
	  int totalRowCount= xlutil.getRowcount("Sheet1");
	  int totalcolumnCount = xlutil.getColumncount("Sheet1",1);
	  
	    Object[][] fields = new Object[totalRowCount][totalcolumnCount];
	  
	  for(int i=1;i<=totalRowCount;i++)
		{
		  for (int j =0; j <totalcolumnCount; j++) {
			  
			  fields[i][j]= xlutil.getCellValue("Sheet1",i,j);
			  
		  }
		
	  }
	 return fields;
	  
  }
  
  
  
   
}
