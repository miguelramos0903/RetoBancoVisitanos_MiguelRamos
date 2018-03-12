package text;

import java.io.IOException;

import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class Visitanos {

	public static void main(String[] args) throws IOException, InterruptedException {
		// TODO Auto-generated method stub
		
		int NumFilas = LeerExcel.ContarFilas();
		System.out.println("**************************");
		System.out.println(NumFilas);
		System.out.println("**************************");
		
		for(int i = 1; i <= NumFilas; i ++) 
    	{		
			String Url = LeerExcel.CargarUrl(i);
			System.out.println(Url);
			System.out.println("**************************");
			
			String Direccion = LeerExcel.CargarDireccion(i);
			System.out.println(Direccion);
			System.out.println("**************************");
        	{
	    		if (Url.equals("N/A") == true || Url.equals("") == true)
	    		{
			                				
	    		}else {
	    			System.setProperty("webdriver.chrome.driver", "C:\\Users\\mramos\\eclipse\\Eclipse-Oxigeno\\eclipse\\ProgramasJava\\Reto_Banco_Visitanos\\Files\\Driver_Browser\\chromedriver.exe");
			    	//System.setProperty("webdriver.gecko.driver", "C:\\Users\\mramos\\eclipse\\Eclipse-Oxigeno\\eclipse\\ProgramasJava\\Reto_Banco_Visitanos\\Files\\Driver_Browser\\geckodriver.exe");
			        WebDriver driver = new ChromeDriver();
			        driver.get(Url);
			        Thread.sleep(3000);
			        
			        driver.manage().window().maximize();
			        Thread.sleep(5000);
			        driver.findElement(By.xpath("//*[@id=\"footer-content\"]/div[1]/div/div/div[4]/div/a/img")).click();
			        Thread.sleep(3000);
			        driver.findElement(By.id("srch-term")).click();
			        driver.findElement(By.id("srch-term")).clear();
			        driver.findElement(By.id("srch-term")).sendKeys(Direccion);
			        Thread.sleep(3000);
			        driver.findElement(By.id("srch-term")).sendKeys(Keys.ENTER);
			        Thread.sleep(5000);
			        
			        System.out.println(driver.findElement(By.xpath("//*[@id=\'tab1\']/div[1]/div[6]/div[1]")).getText());
			        System.out.println("**************************");
			        
			        driver.findElement(By.xpath("//*[@id='tab1']/div[1]/div[6]/div[" + 1 + "]/div/div[1]/button")).click();
			        
			        String TextObtenido = driver.findElement(By.xpath("//*[@id=\'tab1\']/div[1]/div[6]/div[1]")).getText();
			        LeerExcel.EscribirExcel(TextObtenido, i);
			        
			        driver.close();
			        driver.quit();
	    		}
        	}
    	}
	}

}
