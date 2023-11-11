/**
 * 
 */
package org.project.encryptExcel;

import java.io.FileInputStream;
import java.io.IOException;
import java.security.GeneralSecurityException;
import java.util.Properties;

import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.junit.Assert;
import org.junit.Test;
import org.project.encryptExcel.utils.PropertyUtil;

/**
 * 
 */
public class EncryptExcelTest {
	
	private Properties props = PropertyUtil.getProps();
	private String outputPath = props.getProperty("outputPath");
	
	@Test
	public void test() {
		try {
			POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(outputPath+"test1.xlsx"));
			EncryptionInfo info = new EncryptionInfo(fs);
			
			Assert.assertTrue(info.getDecryptor().verifyPassword("1234abc"));
			
		} catch (GeneralSecurityException | IOException e) {
			e.printStackTrace();
		}
	}

}
