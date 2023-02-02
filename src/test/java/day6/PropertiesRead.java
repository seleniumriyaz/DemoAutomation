package day6;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Properties;

public class PropertiesRead {

	public static void main(String[] args) throws IOException {
		
		FileInputStream fis1=new FileInputStream("src\\test\\resources\\properties\\config.properties");
		
		Properties pr=new Properties();
		
		pr.load(fis1);
		
		System.out.println(pr.getProperty("browser"));
		

	}

}
