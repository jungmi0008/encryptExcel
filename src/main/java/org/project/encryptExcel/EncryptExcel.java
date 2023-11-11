package org.project.encryptExcel;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.OutputStream;
import java.net.InetAddress;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;

import org.apache.poi.hssf.record.crypto.Biff8EncryptionKey;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.project.encryptExcel.utils.PropertyUtil;

public class EncryptExcel 
{	
	private static Properties props = PropertyUtil.getProps();
	private static String targetFileList = props.getProperty("targetFileList");
	private static String inputPath = props.getProperty("inputPath");
	private static String outputPath = props.getProperty("outputPath");
	private static String failurePath = props.getProperty("failurePath");

	private static int errorCnt = 0;
	
	/**
	 * create file list
	 * 
	 * @param targetFile
	 * @param password
	 * @throws Exception
	 */
	@SuppressWarnings("resource")
	private static List<File> getExcelFileList() throws Exception{
		
		List<File> fileList = new ArrayList<File>();
		
		try {
			//1.Reading targetFileList.txt file using BufferedReader
			FileReader fr = new FileReader(new File(targetFileList));
			BufferedReader reader = new BufferedReader(fr);
			
			String sLine = null;
			
			//2.Creating target file list
			while((sLine = reader.readLine()) != null ) {
				File targetFile = new File(inputPath+sLine);
				if(!targetFile.exists()){
					throw new Exception("There's no file name : "+sLine+". Please check targetFileList.txt.");
				}
				fileList.add(targetFile);
			}
			
			if(fileList.size() == 0) {
				throw new Exception("There's no file name. Please check targetFileList.txt.");
			}
			
			reader.close();
			
		} catch (Exception e) {
			throw e;
		}
		
		return fileList;
	}
	
	/**
	 * encrypt for xls(Binary formats) file
	 * 
	 * @param targetFile
	 * @param password
	 * @throws Exception
	 */
	private static void encryptExcelForHSSF(File targetFile, String password){
		
		try {
			//1.Setting password
	        Biff8EncryptionKey.setCurrentUserPassword(password);

	        //2.Creating workbook
	        FileInputStream fis =  new FileInputStream(targetFile);
	        HSSFWorkbook wb = new HSSFWorkbook(fis);
	        
	        LocalDate now = LocalDate.now();
	        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyyMMdd");
	        String formatedNow = now.format(formatter);
	        
	        String finalPath = outputPath+formatedNow+"_"+targetFile.getName();
	        if(targetFile.getName().contains("19")) {
	        	finalPath = failurePath+formatedNow+"_"+targetFile.getName();
	        }
	        
	        //3.Creating file output stream to write the workbook data
	        FileOutputStream fos = new FileOutputStream(finalPath);

	        //4.Writing workbook on output stream
	        wb.write(fos);

	        fos.close();
	        wb.close();
		    
	        //5.Creating the password text file
	        savePassword(targetFile, password);
	        
	        insertLog(targetFile, new File(finalPath));
	        
	    } catch (Exception e) {
	    	errorCnt++;
	        e.printStackTrace();
	    }
	}
	
	/**
	 * encrypt for xlsx(XML-based formats) file
	 * 
	 * @param targetFile
	 * @param password
	 * @throws Exception
	 */
	private static void encryptExcelForXSSF(File targetFile, String password) {
		try {
    	    //1.Setting encryption information
	        POIFSFileSystem fs = new POIFSFileSystem();
	        EncryptionInfo encInfo = new EncryptionInfo(EncryptionMode.agile);
	        
	        //2.Setting password
	        Encryptor enc = encInfo.getEncryptor();
	        enc.confirmPassword(password);
	 
		    //3.Reading OOXML file and writing to encrypted output stream
	        try (OPCPackage opc = OPCPackage.open(targetFile, PackageAccess.READ_WRITE);
                OutputStream os = enc.getDataStream(fs)) {
                opc.save(os);
                os.close();
            }
	        
	        LocalDate now = LocalDate.now();
	        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyyMMdd");
	        String formatedNow = now.format(formatter);
	        
	        String finalPath = outputPath+formatedNow+"_"+targetFile.getName();
	        if(targetFile.getName().contains("19")) {
	        	finalPath = failurePath+formatedNow+"_"+targetFile.getName();
	        }
	        
	        try (FileOutputStream fos = new FileOutputStream(finalPath)) {
	            fs.writeFilesystem(fos);
	            fos.close();
	            fs.close();
	        }
		    
	        //5.Creating the password text file
	        savePassword(targetFile, password);
	        
	        insertLog(targetFile, new File(finalPath));
	        
		    }catch (Exception e) {
		    	errorCnt++;
		    	e.printStackTrace();
			}
	}
	
	/**
	 * create password list file
	 * 
	 * @param targetFile
	 * @param password
	 * @throws IOException
	 */
	private static void savePassword(File targetFile, String password) {
		try {
			
			//1.Creating the password file
			File pwdFile = new File(outputPath+targetFile.getName()+".pwd.txt");
			
			if(pwdFile.exists()) {
				pwdFile.delete();
			}
			
			//2.Writing the password using BufferedWriter
			FileWriter fw = new FileWriter(pwdFile, true);
			BufferedWriter writer = new BufferedWriter(fw);
			
			writer.write(password);
			
			writer.close();
			
		} catch (IOException e) {
			e.printStackTrace();
		}
		
	}
	
	private static void insertLog(File inputFile, File outputFile) {
		try {
			//4.Writing out the encrypted version to the output directory
	        LocalDate now = LocalDate.now();
	        // 포맷 정의하기
	        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyyMMdd");
	 
	        // 포맷 적용하기
	        String formatedNow = now.format(formatter);
	        
	        String hostName = InetAddress.getLocalHost().getHostName();
	        
			File logFile = new File(outputPath+formatedNow+"_log.txt");
			
			if(!logFile.exists()) {
				logFile.createNewFile();
			}
			
			FileWriter fw = new FileWriter(logFile, true);
			BufferedWriter writer = new BufferedWriter(fw);
			
			writer.write("inputPath : "+inputFile.getAbsolutePath()+"\n");
			writer.write("outputPath : "+outputFile.getAbsolutePath()+"\n");
			writer.write("original file name : "+inputFile.getName()+"\n");
			writer.write("encrypted file name : "+outputFile.getName()+"\n");
			writer.write("owned by : "+hostName+"\n\n");
			
			writer.close();
			
		} catch (IOException e) {
			e.printStackTrace();
		}
		
	}
	
	public static void main(String[] args) {
		
		System.out.println("start encryption process");
		
		List<File> fileList = null;
		
		try {
			//1.Getting excel file list
			fileList = new ArrayList<File>();
			fileList = getExcelFileList();
			
			//2.Setting password
			String password = "1234abc";
			
			//3.checking the extension of excel file and choosing the process to encrypt each excel file
			for (File targetFile : fileList) {
				String fileName = targetFile.getName();
				String fileExtension = fileName.substring(fileName.lastIndexOf(".") + 1);
				
				System.out.println("target file name : "+fileName);
				//Binary formats
				if(fileExtension.equals("xls")) {
					encryptExcelForHSSF(targetFile, password);
				
				//XML-based formats
				} else if (fileExtension.equals("xlsx")){
					encryptExcelForXSSF(targetFile, password);
				}
			}
			
		} catch (Exception e) {
			e.printStackTrace();
			
		} finally {
			System.out.println("Success Count : "+(fileList.size()-errorCnt)+", errCount : "+errorCnt);
			System.out.println("end encrypttion process");
		}
	}

}
