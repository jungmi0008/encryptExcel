package org.project.encryptExcel.utils;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Properties;

public class PropertyUtil {
	
	private static String configPath = "C:\\Users\\jungmi\\eclipse-workspace\\encryptExcelProject\\encryptExcel\\config.properties";
	
	private static Properties props = new Properties();
	
	public static Properties getProps() {
		
		try {
			props.load(new FileInputStream(configPath));
			
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		return props;
	}
}