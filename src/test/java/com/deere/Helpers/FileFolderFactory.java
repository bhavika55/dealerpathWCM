
/* 
 * Project    : DealerPath
 * Script     : FileFolderFactory
 * Author     : Shrishail Baddi
 * Date       : April.08.2018
 * Last Modified On:
 * Modified By :
 */



package com.deere.Helpers;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Properties;
import java.util.regex.Pattern;

import org.testng.Assert;

public class FileFolderFactory {

	// This method will create the folder with name as parameter @strNewFolderName
	// in the working directory

	public static void createDirectory(String strNewFolderName) {

		try {

			if (isValidName(strNewFolderName)) {
				
				File dir = new File(BaseClass.strWorkingDir + strNewFolderName);
				if (!dir.exists()) {
					dir.mkdir();
					LogFactory.info("Folder " + strNewFolderName + " created successfully...");
				} else {
						LogFactory.info("Unable to create the folder " + strNewFolderName + " as it's already Exists");
				}

			}
		} catch (Exception e) {
			LogFactory.error(e.getMessage());
		}

	}

	// This method will create the folder with name as @strNewFolderName in the user
	// specified path @strFolderPath

	public static String createDirectory(String strFolderPath, String strNewFolderName) {

		  try {
		   
		    if (isValidName(strNewFolderName)) {
		     
		    File dir = new File(strFolderPath + strNewFolderName);
		    if (!dir.exists()) {
		     dir.mkdir();
		     
		     LogFactory.info("Folder " + strNewFolderName + " created successfully on " + strFolderPath);
		    }
		    else {
		      LogFactory.info("Unable to create the folder " + strNewFolderName + " as it's already Exists");
		    }
		    return strFolderPath + strNewFolderName;
		   } 

		  } catch (Exception e) {
		   LogFactory.error(e.getMessage());
		  }
		  return BaseClass.strScreenshotDir + strNewFolderName;

		 }

	public static HashMap<String, String> getFileNameAndExtension(File fileFileName) {
		String fileName = fileFileName.getName();
	
		if (isValidName(fileName))
			return (getFileNameAndExtension(fileName));
		else
			return null;

	}

	
	
		public static HashMap<String, String>  getFileNameAndExtension(String strFileName) {
			
			
			HashMap<String, String> fileDetail = new HashMap <String, String> ();

			
		if (isValidName(strFileName)) {
		
		
				if (strFileName.lastIndexOf(".") != -1 && strFileName.lastIndexOf(".") != 0) {
					int intFromIndex = strFileName.lastIndexOf('.');

					fileDetail.put("fileName", strFileName.substring(0,intFromIndex));
					fileDetail.put("extension", strFileName.substring(intFromIndex));
				
					
				}
				return fileDetail; 
		}
		else {
				LogFactory.info("Invalid File");
				return null;
			 }

		
		}
		
	public static String getExtension(File file) {
		if (file != null) {
			final String path = file.getAbsolutePath();
			try {
				return path.substring(path.lastIndexOf('.') + 1).toLowerCase();
			} catch (Exception ex) {
				return "_blank";
			}
		} else {
			return "_blank";
		}
	}

	/**
	 * Returns the name of a file if it exists. Returns null if the file is missing.
	 * 
	 * @param filePath
	 *            The file path. Should end with a \ or /
	 * @param fileName
	 *            Name of the file (for example: image.jpg)
	 */

	public String findIfExisting(String filePath, String fileName) {
		
		File f = new File(filePath + fileName); // Create path to file
		if (f.exists()) {
			// If the file exists, the file name is returned.
			return fileName;
		} else {
			// If file is missing, returns nothing.
			return null;
		}
	}

	
	// this function returns true if the @text is valid filename

	public static boolean isValidName(String text) {
		
		Pattern pattern = Pattern.compile("[//\\:*?\"<>|]");
		if (pattern.matcher(text).find())
			
			return false;
		else
			return true;
	}

	
	// this method will return the file extension and file name from the valid URL's like
	// url="http://orig02.deviantart.net/9d0c/f/2013/044/a/2/luffy_faces_by_constanza_chan14-d5utbf9.png";
	// strFileName= luffy_faces_by_constanza_chan14-d5utbf9.png;
	
	// returns the filename and extension of file
	
	public static HashMap<String, String>  getFileNameFromURL(String strFileURL) {
		
		String strFileName= strFileURL.substring(strFileURL.lastIndexOf("/") + 1, strFileURL.length());
		
		if (isValidName(strFileName))
			return (getFileNameAndExtension(strFileName));
		else
			return null;
	
	}

	
	
	/**
	 * 
	 * @param fileName
	 *            - expects the name of the file to be created, without extension
	 *            Creates file in the current project directory
	 * 
	 * @return name of the file created
	 * 
	 */
	public static File createExcelFile(String strFilePath, String strFileName) throws FileNotFoundException {
		
		if (isValidName(strFileName)){
			
			String file = strFileName + ".xlsx";
		
			File XLfile = new File (strFilePath + File.separator + file);
		
			if (!XLfile.exists()) {
			
				new FileOutputStream(XLfile);
				return XLfile;
				
			} 
			else { 
					LogFactory.info("File " + strFileName + ".xlsx" + " already exists in the path " + strFilePath);
					return XLfile;
			}
		}
		else {
			   LogFactory.info("Unable to create the file " + strFileName + ".xlsx" + " as filename is invalid");
			   return null;
		}
	}
	

	
	
	/**
	 * 
	 * @param fileName
	 *            - expects the name of the file to be created, without extension
	 *            Appends timestamp and creates file <fileName>_<timeStamp>.xlsx
	 *            Creates file in the current project directory
	 * 
	 * @return name of the file created
	 * 
	 */
	
	
	
	public static File createExcelFileWithTimeStamp(String strFilePath, String strFileName) throws FileNotFoundException {
		
		if (isValidName(strFileName)){
			
			String date = new SimpleDateFormat("yyyyMMddHHmm").format(new Date());
			String file = strFileName + "_" + date + ".xlsx";
		
			File XLfile = new File (strFilePath + File.separator + file);
		
			if (!XLfile.exists()) {
			
				new FileOutputStream(XLfile);
				return XLfile;
				
			} 
			else { 
					LogFactory.info("File " + strFileName + ".xlsx" + " already exists in the path " + strFilePath);
					return XLfile;
			}
		}
		else {
			   LogFactory.info("Unable to create the file " + strFileName + ".xlsx" + " as filename is invalid");
			   return null;
		}
	}

	
	/**This method is used to load the Configuration file.
	 * 
	 * @ @param The name of the configuration file
	 * @return Properties it will return the instance of property file after loading
	 * @exception IOException: If there is any Input/Output Exception is found
	 */
	public static Properties loadConfigFile(String loadController){
		
		try{
			BaseClass.propConfig = new Properties();
			if(isFilePresent(BaseClass.strWorkingDir+"\\src\\test\\java\\com\\deere\\ConfigFactory\\"+loadController)){
				
				FileInputStream fp = new FileInputStream(BaseClass.strWorkingDir+"\\src\\test\\java\\com\\deere\\ConfigFactory\\"+ loadController);
				BaseClass.propConfig.load(fp);
		
				LogFactory.info("The configuration file is successfully loaded");
				
			}else{
				LogFactory.info("The configuration file loading is failed");
				
			}
			
		}catch(IOException strInputOutputException){
			
			LogFactory.info("IOException occured in loading the Configuration file");
			Assert.fail("IOException occured in loading the Configuration file");
			
		}catch(Exception e){
			
			e.printStackTrace();
	
			LogFactory.info("Some Exception occured in loading the Configuration file");
			Assert.fail("Some Exception occured in loading the Configuration file");
		}
		
		return BaseClass.propConfig;
	}

	
	
	/**This method is used to check whether the File is present on the path passed as parameter.
	 * 
	 * 
	 * @param key The key for which value needs to be fetched
	 * @param strObjLabel The Object identifier or label displayed in the application
	 * @return True/False (True if the file is present on the path, false otherwise)
	 * @exception Exception: If there is any exception occurs
	 */
	public static boolean isFilePresent(String strFilePath){
		
		try{
			if((strFilePath).trim() == ""){
				
				LogFactory.info("The passed file path paramenter is blank");
				return false;
				
			}else{
				
				File file=new File(strFilePath);
				  boolean exists = file.exists();
				  
				  if(exists) {
					  
					  return true;
					  
				  }else{
					  
					  return false;
				  }
			}
			
		}catch(Exception e){
			
			e.printStackTrace();
			return false;
			
		}
	}
	
	
}
