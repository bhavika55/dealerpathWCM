
/* 
 * Project    : DealerPath
 * Script     : Excel Factory
 * Author     : Shrishail Baddi
 * Date       : April.05.2018
 * Last Modified On:
 * Modified By :
 */

package com.deere.Helpers;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class ExcelFactory extends BaseClass {

	public static XSSFWorkbook book;
	private static XSSFSheet dataSheet;
	
	private static XSSFSheet dataDealerInfoSheet;
	private static XSSFSheet dataWCMContentSheet;
	private static XSSFSheet dataAdditionalTCSheet;
	private static XSSFSheet dataTranslationsSheet;
	
	private static XSSFRow row;
	private static XSSFCell cell;
	static List<LinkedHashMap> dealerinfo;

	/**
	 * @param file
	 *            - expects file name with absolute path
	 * @param sheetName
	 *            - expects name of the sheet to be read
	 * @return instance of XSSFSheet
	 */
	public static XSSFSheet readExcel(String strFile, String strSheetName) {
		try {
			FileInputStream excelFile = new FileInputStream(new File(strFile));
			book = new XSSFWorkbook(excelFile);
		} catch (FileNotFoundException fnfe) {
			System.out.println(fnfe.getMessage());
		} catch (IOException ioe) {
			ioe.getMessage();
		}
		return dataSheet = book.getSheet(strSheetName);
	}

	/**
	 * 
	 * @param sheet
	 *            - expects instance of XSSFSheet
	 * @return number of rows in the sheet
	 */
	public static int getRowCount(XSSFSheet sheet) {
		return sheet.getLastRowNum();
	}

	/**
	 * 
	 * @param sheet
	 *            - expects instance of XSSFSheet
	 * @param rowNum
	 *            - expects the row number to be read
	 * @return data in the row as an instance of XSSFRow
	 */
	public static XSSFRow readRow(XSSFSheet sheet, int rowNum) {
		return sheet.getRow(rowNum);
	}

	/**
	 * 
	 * @param row
	 *            - expects row as instance of XSSFRow
	 * @return number of columns in the given row
	 */
	public static int getColumnCount(XSSFRow row) {
		return row.getPhysicalNumberOfCells();
	}

	/**
	 * 
	 * @param row
	 *            - expects row to be read as an instance of XSSFRow
	 * @return value of each column is read as an object in an ArrayList. Assumes
	 *         that excel contains either numeric or string values
	 */
	public static ArrayList<Object> readRowData(XSSFRow row) {
		int columnCount = 0;
		if (row != null) {
			columnCount = getColumnCount(row);
		}
		ArrayList<Object> rowData = new ArrayList();
		for (int i = 0; i < columnCount; i++) {
			XSSFCell cell = row.getCell(i);
			if (cell.getCellTypeEnum() == CellType.STRING) {
				// System.out.println(cell.getStringCellValue());
				rowData.add(cell.getStringCellValue());
			}
			if (cell.getCellTypeEnum() == CellType.NUMERIC) {

				if (DateUtil.isCellDateFormatted(cell)) {
					rowData.add(new SimpleDateFormat("MM_dd_yyyy").format(cell.getDateCellValue()));
				} else {
					rowData.add(cell.getNumericCellValue());
				}
			}
			if (cell.getCellTypeEnum() == CellType._NONE) {
				rowData.add(null);
			}
			if (cell.getCellTypeEnum() == CellType.BLANK) {
				rowData.add("");
			}
		}
		// System.out.println("Number of values in list: " + rowData.size());
		return rowData;

	}

	/**
	 * 
	 * @param sheet
	 *            - Expects the sheet to be read and an instance of XSSFSheet
	 * @return All the rows of sheet as an ArrayList of Maps. Each row of data is
	 *         converted into HashMap, with all the row headers as keys and cell
	 *         values in rows as values
	 */
	public static ArrayList<Map<String, Object>> readSheetAsList(XSSFSheet sheet) {
		ArrayList<Map<String, Object>> sheetData = new ArrayList();
		int rowCount = sheet.getLastRowNum();
		if (rowCount > 0) {
			ArrayList<String> header = new ArrayList();
			row = sheet.getRow(0);
			int colCount = row.getPhysicalNumberOfCells();
			for (int i = 0; i < colCount; i++) {
				cell = row.getCell(i);
				if (cell.getCellTypeEnum() == CellType.STRING) {
					header.add(cell.getStringCellValue());
					// System.out.println(cell.getStringCellValue());
				}
			}

			System.out.println(header);

			for (int i = 1; i <= rowCount; i++) {
				Map<String, Object> map = new HashMap<>();
				row = sheet.getRow(i);
				for (int j = 0; j < header.size(); j++) {
					cell = row.getCell(j);
					if (cell == null) {
						map.put(header.get(j), null);
					} else {
						if (cell.getCellTypeEnum() == CellType._NONE) {
							map.put(header.get(j), null);
						}
						if (cell.getCellTypeEnum() == CellType.STRING) {
							map.put(header.get(j), cell.getStringCellValue());

						}
						if (cell.getCellTypeEnum() == CellType.NUMERIC) {

							if (DateUtil.isCellDateFormatted(cell)) {
								map.put(header.get(j),
										new SimpleDateFormat("MM_dd_yyyy").format(cell.getDateCellValue()));
							} else {
								map.put(header.get(j), cell.getNumericCellValue());
							}
						}

						if (cell.getCellTypeEnum() == CellType.BLANK) {
							map.put(header.get(j), "");
						}
					}

				}

				sheetData.add(map);
			}

		}
		System.out.println(sheetData);
		return sheetData;
	}

	
/*	public static HashMap<String, String> getUserAddtionalTestcases() throws IOException, Exception {
		try {

			XSSFRow singleRow;
			XSSFCell cell;
			int TestCase_index = 0;
			int TestData_index = 0;

			LinkedHashMap<String, String> Mapobj = new LinkedHashMap<String, String>();

			// Excel Read
			 dataSheet = book.getSheet("Additional Testcases");

			int totalRows = dataSheet.getLastRowNum();
			int totalCol = dataSheet.getRow(4).getLastCellNum();
			for (int j = 0; j < totalCol; j++) {

				String Attribute_name = dataSheet.getRow(4).getCell(j).getStringCellValue().trim();
				// Passing Preferred Language Column name
				if (Attribute_name.equalsIgnoreCase(strUserRACFID)) {
					TestData_index = j;
					break;
				}
				// Passing Test Case Column name
				else if (Attribute_name.equalsIgnoreCase(strTestCase)) {
					TestCase_index = j;
				}
			}

			for (int i = 1; i <= totalRows; i++) {
				singleRow = dataSheet.getRow(i);
				
				if ((singleRow.getCell(TestCase_index).getCellType() == Cell.CELL_TYPE_STRING)) {
					
					if ((singleRow.getCell(TestCase_index).getStringCellValue().trim()).length() > 0) {
						
						String Testcase_id = singleRow.getCell(TestCase_index).getStringCellValue().trim();
						if (!Testcase_id.equalsIgnoreCase("NA")) {
							// System.out.println(Testcase_id);
							cell = singleRow.getCell(TestData_index);
							int cellvalue_data = 0;

							if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
								Mapobj.put(Testcase_id, cell.getStringCellValue());
								//System.out.println(Testcase_id + Mapobj.get(Testcase_id));
								} else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
								cellvalue_data = (int) cell.getNumericCellValue();
								String cellvalue_String = Integer.toString(cellvalue_data);
								Mapobj.put(Testcase_id, cellvalue_String);
								//System.out.println(Testcase_id + Mapobj.get(Testcase_id));
								} else if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
								String cellvalue_String = cell.getStringCellValue();
								Mapobj.put(Testcase_id, cellvalue_String);
								//System.out.println(Testcase_id + Mapobj.get(Testcase_id));
								} else if (cell.getCellType() == Cell.CELL_TYPE_BLANK) {
									Mapobj.put(Testcase_id, "");
							//System.out.println(Testcase_id + Mapobj.get(Testcase_id));
								} else {
									Mapobj.put(Testcase_id, "");
									//System.out.println(Testcase_id + Mapobj.get(Testcase_id));
								}

						}
					}	
				}

			}
	//		System.out.println(Mapobj);
			 
			return Mapobj;

		} catch (Exception e) {
			LogFactory.info(e.getMessage());
			return null;
		}

	}
*/
	
	public static HashMap<String, String> getUserAddtionalTestcases() throws IOException, Exception {
	try {

		XSSFRow singleRow;
		XSSFCell cell;
		int TestRacfData_index = 0;
		int TestCaseID_Index = 0;
		int TestEnglishData_index = 0;
		
		LinkedHashMap<String, String> Mapobj = new LinkedHashMap<String, String>();
		// Excel Read
		int totalRows = dataAdditionalTCSheet.getLastRowNum();
		int totalCol = dataAdditionalTCSheet.getRow(4).getLastCellNum();

		for (int j = 0; j < totalCol; j++) {
			String Attribute_name = dataAdditionalTCSheet.getRow(4).getCell(j).getStringCellValue().trim();
			// Passing Preferred Language Column name
			
			if (Attribute_name.equalsIgnoreCase(strUserRACFID)) {
				TestRacfData_index = j;
				break;
			}
			else if (Attribute_name.equalsIgnoreCase("English")) {
				TestEnglishData_index = j;
			}			
			// Passing Test Case Column name
			else if (Attribute_name.equalsIgnoreCase(strTestCase)) {
				TestCaseID_Index = j;
			}
		}

		for (int i = 1; i <= totalRows; i++) {
			singleRow = dataAdditionalTCSheet.getRow(i);
			if ((singleRow.getCell(TestCaseID_Index).getCellType() == Cell.CELL_TYPE_STRING)) {
				
				if ((singleRow.getCell(TestCaseID_Index).getStringCellValue().trim()).length() > 0) {
					String Testcase_id = singleRow.getCell(TestCaseID_Index).getStringCellValue().trim();
					
					if (!Testcase_id.equalsIgnoreCase("NA")) {
						cell = singleRow.getCell(TestRacfData_index);

						if (cell.getStringCellValue().toString().trim().isEmpty() || cell.getStringCellValue().toString().trim()== null || cell.getCellType() == Cell.CELL_TYPE_BLANK) {
							cell = singleRow.getCell(TestEnglishData_index);
						}
						
						int cellvalue_data = 0;
						if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
							Mapobj.put(Testcase_id, cell.getStringCellValue());
							} else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
								
							cellvalue_data = (int) cell.getNumericCellValue();
							String cellvalue_String = Integer.toString(cellvalue_data);
							Mapobj.put(Testcase_id, cellvalue_String);
							//System.out.println(Testcase_id + Mapobj.get(Testcase_id));
						
							} else if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
							String cellvalue_String = cell.getStringCellValue();
							Mapobj.put(Testcase_id, cellvalue_String);
							//System.out.println(Testcase_id + Mapobj.get(Testcase_id));
							} else if (cell.getCellType() == Cell.CELL_TYPE_BLANK) {
								Mapobj.put(Testcase_id, "");
								//System.out.println(Testcase_id + Mapobj.get(Testcase_id));
							} else {
								Mapobj.put(Testcase_id, "");
								//System.out.println(Testcase_id + Mapobj.get(Testcase_id));
							}

					}
				}	
			}

		}
		System.out.println(Mapobj);
		 
		return Mapobj;

	} catch (Exception e) {
		LogFactory.info(e.getMessage());
		return null;
	}

}

	
	
	
	/*
	 * public static HashMap<String, String> translationLookUp(String tarnsCode)
	 * throws IOException, Exception {
	 */
	@Test
	public static HashMap<String, String> translationLookUp(String tarnsCode) throws IOException, Exception {
		try {

			// String tarnsCode = "fr_FR";
			XSSFRow singleRow;
			XSSFCell cell;
			int intBaseLangCodeIndex = 0;
			int intTransCodeIndex = 0;

			LinkedHashMap<String, String> Mapobj = new LinkedHashMap<String, String>();
			/*
			 * FileInputStream input = new FileInputStream(
			 * "C:\\Users\\shrishail.baddi\\Desktop\\DealerPath\\Shrey\\DealerPath_6June\\TestData\\testdatasample.xlsx"
			 * );
			 */
			// book = new XSSFWorkbook(input);
			// Excel Read
			
			int totalRows = dataTranslationsSheet.getLastRowNum();
			int totalCol = dataTranslationsSheet.getRow(4).getLastCellNum();

			for (int j = 0; j < totalCol; j++) {

				String Language_Code = dataTranslationsSheet.getRow(1).getCell(j).getStringCellValue().trim();

				if (Language_Code.equalsIgnoreCase(tarnsCode)) {
					intTransCodeIndex = j;
					if (intTransCodeIndex == 0)
						intBaseLangCodeIndex = 0;
					break;
				}

				else if (Language_Code.equalsIgnoreCase("en_US")) {
					intBaseLangCodeIndex = j;
				}
			}

			for (int i = 2; i <= totalRows; i++) {

				singleRow = dataTranslationsSheet.getRow(i);
				if ((singleRow.getCell(intBaseLangCodeIndex).getCellType() == Cell.CELL_TYPE_STRING)) {

					if ((singleRow.getCell(intBaseLangCodeIndex).getStringCellValue().trim()).length() > 0) {
						String Testcase_id = singleRow.getCell(intBaseLangCodeIndex).getStringCellValue().trim();
						// System.out.println(Testcase_id);
						cell = singleRow.getCell(intTransCodeIndex);
						String strCellValue = cell.toString();
						int cellvalue_data = 0;

						if (cell.getCellType() == Cell.CELL_TYPE_STRING) {

							if (!strCellValue.equalsIgnoreCase("X")) {
									Mapobj.put(Testcase_id, cell.getStringCellValue());
									// System.out.println(Testcase_id +Mapobj.get(Testcase_id));
							} else {
									Mapobj.put(Testcase_id, Testcase_id);
									// System.out.println(Testcase_id +Mapobj.get(Testcase_id));
							}

						} else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
							cellvalue_data = (int) cell.getNumericCellValue();
							String cellvalue_String = Integer.toString(cellvalue_data);
							Mapobj.put(Testcase_id, cellvalue_String);
							// System.out.println(Testcase_id + Mapobj.get(Testcase_id));

						} else if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
							String cellvalue_String = cell.getStringCellValue();
							Mapobj.put(Testcase_id, cellvalue_String);
							// System.out.println(Testcase_id + Mapobj.get(Testcase_id));
						} else if (cell.getCellType() == Cell.CELL_TYPE_BLANK) {
							Mapobj.put(Testcase_id, "");
							// System.out.println(Testcase_id + Mapobj.get(Testcase_id));

						} else {
								Mapobj.put(Testcase_id, "");
								// System.out.println(Testcase_id + Mapobj.get(Testcase_id));
						}

					}
				}

			}
			// System.out.println(Mapobj);
			return Mapobj;

		} catch (Exception e) {
			LogFactory.info(e.getMessage());
			return null;
		}
	}

	/*
	 * public static void main(String[] args) throws Exception, Exception {
	 * HashMap<String, String> obj = ReadingExcelData_as_preffered_language(
	 * "D:\\Workspace\\DealerPath\\DealerPath-master\\DealerPath\\TestData\\TestData.xlsx",
	 * "Admin"); System.out.println(obj);
	 * System.out.println(obj.get("TC02_Homepage")); }
	 */

	/**
	 * 
	 * @param fileName
	 *            - expects absolute path of the file, to which data is to be
	 *            written
	 * @param rowList
	 *            - expects the data to be written in the row Data will be written
	 *            in the row after the last existing row
	 * 
	 */
	public static void writeRow(String fileName, List<Map<String, String>> rowList)
			throws IOException, InvalidFormatException {
		File oFile = new File(fileName);
		FileInputStream input = new FileInputStream(oFile);
		book = new XSSFWorkbook(input);
		if(fileName.contains("WCM"))
		{
			dataSheet = book.getSheet("WCM Content");	
		}
		else
		{
			dataSheet = book.getSheet("TestResults");
			
		}
		
		int rowNum = dataSheet.getLastRowNum();
		String cellVal;
		CellStyle style = book.createCellStyle();// *
		Font font = book.createFont();// *

		for (int i = 0; i < rowList.size(); i++) {
			Map<String, String> map = new HashMap<String, String>();
			map = rowList.get(i);
			int cellNumber = 0;
			XSSFRow row = dataSheet.createRow(++rowNum);
			for (Map.Entry<String, String> entry : map.entrySet()) {
				cellVal = entry.getValue();
				// System.out.println("Current Cell: " + cellNumber);
				XSSFCell cell = row.createCell(cellNumber);
				// System.out.println(cell + "-" + cellNumber);

				if (cellVal instanceof String) {
					cell.setCellValue((String) cellVal);
				} else {
					cell.setCellValue("String");
				}

				// dataSheet.autoSizeColumn(cellNumber);
				dataSheet.setColumnWidth(0, 5000);
				dataSheet.setColumnWidth(1, 13000);
				dataSheet.setColumnWidth(2, 5000);
				dataSheet.setColumnWidth(3, 13000);
				dataSheet.setColumnWidth(4, 13000);
				dataSheet.setColumnWidth(5, 5000);
				style.setWrapText(true);
				cell.setCellStyle(style);

				cellNumber++;
			}
		}
		input.close();
		FileOutputStream fos = new FileOutputStream(oFile);
		book.write(fos);
		book.close();
		fos.close();
	}

	/**
	 * 
	 * @param fileName
	 *            - expects absolute path of the file, to which data is to be
	 *            written
	 * @param headerList
	 *            - expects all the column names in the file, as a List of String
	 * @return name of the file with absolute path, in which the header is written.
	 */

	public static String writeHeader(String fileName, List<String> headerList) throws IOException {
		FileOutputStream fos = new FileOutputStream(new File(fileName));
		XSSFWorkbook book = new XSSFWorkbook();
		
		XSSFSheet sheet;
		
		sheet = book.createSheet("TestResults");
				
		Row row = sheet.createRow(0);

		int cellNumber = 0;
		Font font = book.createFont();
		font.setBold(true);
		font.setFontHeightInPoints((short) 14);
		CellStyle cellStyle1 = book.createCellStyle();

		for (String header : headerList) {
			Cell cell = row.createCell(cellNumber++);
			cellStyle1.setFont(font);
			cellStyle1.setFillBackgroundColor(IndexedColors.LIGHT_GREEN.getIndex());
			cell.setCellStyle(cellStyle1);
			cell.setCellValue(header);
			sheet.autoSizeColumn(cellNumber);

		}
		book.write(fos);
		book.close();
		fos.close();
		return fileName;
	}

	/*
	 * public static void setAddtionalTestcasesDataMap() throws IOException,
	 * Throwable {
	 * 
	 * LogFactory.info("----------- Reading Excel Test Data Sheet -----------------"
	 * ); LogFactory.info("----------  Selectd sheet Name : " +
	 * BaseClass.strUserRole);
	 * 
	 * BaseClass.mapAddtionalTestcase =
	 * getUserAddtionalTestcases(BaseClass.strDataPath, BaseClass.strUserRole);
	 * LogFactory.info("----------  Selectd  Preffered Language for test data : " +
	 * BaseClass.strUserPrefLang);
	 * 
	 * LogFactory.info("----------  Addtional Testcases Sheet Test Data are : " +
	 * BaseClass.mapAddtionalTestcase);
	 * 
	 * }
	 */
	public static List<LinkedHashMap> getWCMContentDetails(String ContentType) {
		LinkedHashMap rowdata = null;
		List<LinkedHashMap> rowDataList = new ArrayList<LinkedHashMap>();
		for (int i = 0; i < wcmExceldata.size(); i++) {
			rowdata = wcmExceldata.get(i);
			// System.out.println("rowdata"+rowdata);
			String id_value = (String) rowdata.get("Dealer_User_id");
			// System.out.println(id_value);
			if (id_value.equalsIgnoreCase(strUserRACFID)) {

				String value = (String) rowdata.get("ContentType");
				if (value.equalsIgnoreCase(ContentType)) {
					System.out.println(rowdata);
					rowDataList.add(rowdata);
					
				}
			}

		}
		
		return rowDataList;
	}

	public static List<LinkedHashMap> getWCMSiteAreaDetails(String portletlinks) {
		LinkedHashMap rowdata = null;
		List<LinkedHashMap> rowDataList = new ArrayList<LinkedHashMap>();

		LogFactory.info("Reading WCM portlets links page content for " + strUserRACFID);

		for (int i = 0; i < userWCMData.size(); i++) {
			
			rowdata = userWCMData.get(i);
			// System.out.println("rowdata"+rowdata);
			String id_value = (String) rowdata.get("Dealer_User_id");
			// String contentType = (String) rowdata.get("ContentType");
			// System.out.println(id_value);

			if (id_value.equalsIgnoreCase(strUserRACFID)) {

				String portletLinkArea = (String) rowdata.get("3rdLevelIndexPage");
				String contenttype = (String) rowdata.get("ContentType");
				String depertment = (String) rowdata.get("DepartmentName");
				String secondndlevel = (String) rowdata.get("2ndLevel");

				if (portletLinkArea.equalsIgnoreCase(portletlinks) && !contenttype.equalsIgnoreCase("AT_Alerts")
						&& !contenttype.equalsIgnoreCase("AT_Announcement") && !depertment.equalsIgnoreCase("NA")
						&& !secondndlevel.equalsIgnoreCase("NA") && !contenttype.equalsIgnoreCase("AT-Index page")
						&& !contenttype.equalsIgnoreCase("AT-ChildIndex Page")
						&& !contenttype.equalsIgnoreCase("AT-GrandChild Index Page")) {
						// System.out.println(rowdata);
						rowDataList.add(rowdata);
						// break;
				}

			}

		}
		return rowDataList;
	}

	/*
	 * public static LinkedHashSet<String> readExcelData() throws Exception {
	 * FileInputStream inputStream = new FileInputStream(new
	 * File(BaseClass.strDataPath));
	 * 
	 * @SuppressWarnings("resource") XSSFWorkbook workbook = new
	 * XSSFWorkbook(inputStream); XSSFSheet firstSheet =
	 * workbook.getSheet("TestData"); //Exceldata = null; int totalRows =
	 * firstSheet.getLastRowNum();
	 * 
	 * int startRow = 0; int totalCol = 0; //
	 * 
	 * ArrayList<String> ListAttributeName = new ArrayList<String>(); for (int i =
	 * 0; i < totalRows; i++) {
	 * 
	 * String cellvalue =
	 * firstSheet.getRow(i).getCell(0).getStringCellValue().toString().trim();
	 * 
	 * if (cellvalue.equalsIgnoreCase("User_Id")) { startRow = i;
	 * 
	 * totalCol = firstSheet.getRow(startRow).getLastCellNum();
	 * 
	 * for (int j = 0; j < totalCol; j++) {
	 * 
	 * String AttributeName =
	 * firstSheet.getRow(startRow).getCell(j).getStringCellValue().trim()
	 * .toString();
	 * 
	 * ListAttributeName.add(AttributeName); }
	 * 
	 * break;
	 * 
	 * }
	 * 
	 * } // **************** Reading Excel sheet line by line
	 * 
	 * // Exceldata = new ArrayList<LinkedHashMap>();
	 * 
	 * LinkedHashSet<String> UniqueUserIdList = new LinkedHashSet<String>();
	 * 
	 * for (int i = startRow + 1; i <= totalRows; i++) { String User_Id =
	 * firstSheet.getRow(i).getCell(0).getStringCellValue().trim();
	 * UniqueUserIdList.add(User_Id.toUpperCase()); LinkedHashMap<String, String>
	 * excelrowdata = new LinkedHashMap<String, String>(); //
	 * System.out.println("totalCol" + totalCol); for (int j = 0; j < totalCol; j++)
	 * {
	 * 
	 * String StrAttributeName = ListAttributeName.get(j);
	 * 
	 * String StrAttributeValue =
	 * firstSheet.getRow(i).getCell(j).getStringCellValue().trim();
	 * 
	 * excelrowdata.put(StrAttributeName, StrAttributeValue);
	 * 
	 * } Exceldata.add(excelrowdata); } // Completed excel data return
	 * UniqueUserIdList;
	 * 
	 * }
	 */
	/**
	 * This method is used to fetch the login credentials, log and screenshot status
	 * from excel
	 * 
	 * @author neeraja.mantri
	 * @return LinkedHashMap<String, String>
	 * @throws Throwable
	 * @createdAt 05-06-2018
	 * @modifyBy shrey.choudhary
	 * @modifyAt 07-06-2018
	 */
	public static void setCredentials() throws Throwable {

		try {
			@SuppressWarnings("resource")
			LinkedHashMap<String, String> Mapobj = new LinkedHashMap<String, String>();
			for (int j = 0; j <= dataAdditionalTCSheet.getRow(0).getLastCellNum(); j++) {
				for (int i = 0; i <= 3; i++) {
					
					String Attribute_name = (!(dataAdditionalTCSheet.getRow(i).getCell(j) == null || dataAdditionalTCSheet.getRow(i).getCell(j).toString().isEmpty() ||
											   dataAdditionalTCSheet.getRow(i).getCell(j).getCellType() == cell.CELL_TYPE_BLANK) ) 
							? dataAdditionalTCSheet.getRow(i).getCell(j).getStringCellValue()
								: null;
					if (!Attribute_name.toString().trim().isEmpty() && Attribute_name != null) {
						
						Mapobj.put(Attribute_name,
								dataAdditionalTCSheet.getRow(i).getCell(j + 1).getStringCellValue().toString().trim());
					}
				}
				j = j + 3;

			}
		
			
			strUserName = Mapobj.get("Username").toString().trim();
			strPassword = Mapobj.get("Password").toString().trim();
			URL = Mapobj.get("URL").toString().trim();
			strBrowserType = Mapobj.get("Browser").toString().trim();
			String screenshotPath = Mapobj.get("ScreenshotPath").toString().trim();
			strScreenshotDir = screenshotPath.isEmpty() ? strWorkingDir : screenshotPath;
			ENABLE_SCREENSHOT = Mapobj.get("EnableScreenshot").isEmpty() ? ENABLE_SCREENSHOT = "OFF"
					: Mapobj.get("EnableScreenshot");
			dateformat= Mapobj.get("Date Format").toString().trim();
			System.out.println(Mapobj);

		} catch (Exception e) {
			e.getMessage();
		}
		// TODO: handle exception
	}

	public static List<LinkedHashMap> readDealerInfoData() throws Exception {
		try {

			XSSFCell cell;
			int totalRows = dataDealerInfoSheet.getLastRowNum();
			int startRow = 0;
			int totalCol = 0;
			String HeadervalueList = "";
			String useridvalue = "";
			String Coutryvalue = "";
			String flagDealerPrinciple = "";
			String flagExecutionControl = "";
			String flagAddtionalTestcases = "";
			String flagDealerTypeContorl = "";
			
			int flagExecutionControlIndex = 0;
			int flagAddtionalTestcasesIndex = 0;
			int flagDealerPrincipleIndex = 0;
			int flagDealerTypeIndex=0;

			List<LinkedHashMap> dealerRowDataList = null;
			System.out.println(totalRows);

			ArrayList<String> ListAttributeName = new ArrayList<String>();
			ArrayList<String> dealerHeaderList = new ArrayList<String>();

			for (int i = 1; i < totalRows; i++) {

				String cellvalue = dataDealerInfoSheet.getRow(i).getCell(0).getStringCellValue().toString().trim();

				if (cellvalue.equalsIgnoreCase("Userid")) {
					startRow = i;
					// System.out.println(i);
					totalCol = dataDealerInfoSheet.getRow(startRow).getLastCellNum();

					for (int j = 0; j < totalCol; j++) {

						String dealerHeaderName = dataDealerInfoSheet.getRow(startRow).getCell(j).getStringCellValue().trim()
								.toString();

						switch (dealerHeaderName) {

						case "Construction":
							dealerHeaderList.add("Construction" + "/" + dealerHeaderName);
							break;
							
						case "Utility":
							dealerHeaderList.add("Construction" + "/" + dealerHeaderName);
							break;
							
						case "Forestry":
							dealerHeaderList.add("Forestry" + "/" + dealerHeaderName);
							break;
							
						case "CWP":
							dealerHeaderList.add("CWP" + "/" + dealerHeaderName);
							break;
							
						case "Hitachi":
							dealerHeaderList.add("Hitachi" + "/" + dealerHeaderName);
							 break;
							 
						case "Mining":
							dealerHeaderList.add("Hitachi" + "/" + dealerHeaderName);
							break;

						  case "Ag Equipment":
							dealerHeaderList.add("Agriculture" + "/" + dealerHeaderName);
							break;

						case "Sprayers & Nutrient Applicators":
							dealerHeaderList.add("Agriculture" + "/" + dealerHeaderName);
							break;

						case "Scraper and Scraper Tractor":
							dealerHeaderList.add("Agriculture" + "/" + dealerHeaderName);
							break;

						case "Forage Harvester":
							dealerHeaderList.add("Agriculture" + "/" + dealerHeaderName);
							break;

						case "Commercial":
							dealerHeaderList.add("Turf" + "/" + dealerHeaderName);
							break;

						case "Residential":
							dealerHeaderList.add("Turf" + "/" + dealerHeaderName);
							break;
						
						case "Homeowner":
							dealerHeaderList.add("Homeowner" + "/" + dealerHeaderName);
							break;

						case "JDPS Distributor":
							dealerHeaderList.add("JDPS" + "/" + dealerHeaderName);
							break;

						case "Direct OEM":
							dealerHeaderList.add("JDPS" + "/" + dealerHeaderName);
							break;

						case "Dealer Principle":
							flagDealerPrincipleIndex = j;
							break;
						case "Dealer Type (Main/Sub)":
							flagDealerTypeIndex = j;
							break;	
						case "Execution Control":
							flagExecutionControlIndex = j;
							break;
						case "Additional Testcases":
							flagAddtionalTestcasesIndex = j;
							break;
						default:
							dealerHeaderList.add(dealerHeaderName);
						}

					}
				} else {
					break;
				}
			}
			System.out.println(dealerHeaderList);

			dealerRowDataList = new ArrayList<LinkedHashMap>();

			for (int i = 2; i < totalRows; i++) {
				LinkedHashMap<String, String> RowValuedata = new LinkedHashMap<String, String>();

				useridvalue = dataDealerInfoSheet.getRow(i).getCell(0).getStringCellValue().trim();
				Coutryvalue = dataDealerInfoSheet.getRow(i).getCell(2).getStringCellValue().trim();				
				flagDealerPrinciple = dataDealerInfoSheet.getRow(i).getCell(flagDealerPrincipleIndex).getStringCellValue()
						.trim();
				flagDealerTypeContorl = dataDealerInfoSheet.getRow(i).getCell(flagDealerTypeIndex).getStringCellValue()
						.trim();
				flagExecutionControl = dataDealerInfoSheet.getRow(i).getCell(flagExecutionControlIndex).getStringCellValue()
						.trim();
				flagAddtionalTestcases = dataDealerInfoSheet.getRow(i).getCell(flagAddtionalTestcasesIndex).getStringCellValue()
						.trim();

				HeadervalueList = "";

				for (int j = 2; j < totalCol; j++) {

					cell = dataDealerInfoSheet.getRow(i).getCell(j);

					if (cell.getCellTypeEnum() != CellType.BLANK && cell.getCellTypeEnum() != CellType._NONE) {

						String dealerRowValue = dataDealerInfoSheet.getRow(i).getCell(j).getStringCellValue().trim().toString();

						// HeadervalueList= dealerHeaderList.get(j);

						if (dealerRowValue.equalsIgnoreCase("x")) {
							if (HeadervalueList == "") {
								HeadervalueList = dealerHeaderList.get(j);
							} else {

								HeadervalueList = HeadervalueList + "," + dealerHeaderList.get(j);
							}
						}
					}

					/*
					 * if(HeadervalueList.length()>0) {
					 * System.out.println("------>>>>>> "+HeadervalueList);
					 * //HeadervalueList=HeadervalueList.substring(1,HeadervalueList.length());
					 * 
					 * System.out.println("++++++++++>>>> "+HeadervalueList); }
					 */

				}
				if (useridvalue.length() > 0) {

					RowValuedata.put("Dealer_User_id", useridvalue);
					RowValuedata.put("Dealer_Country", Coutryvalue);
					RowValuedata.put("Execution Control", flagExecutionControl);
					RowValuedata.put("Addtional Testcases", flagAddtionalTestcases);
					RowValuedata.put("Dealer Principle", flagDealerPrinciple);
					RowValuedata.put("Dealer Type (Main/Sub)", flagDealerTypeContorl);
					RowValuedata.put("Dealer_ProductType", HeadervalueList);

					dealerRowDataList.add(RowValuedata);

				}

			}
			System.out.println(dealerRowDataList);
			return dealerRowDataList;
		} catch (Exception e) {
			System.out.println(e.getMessage());
			return null;
		}

	}

	public static HashSet<String> uniqueRACFId() throws Throwable {

		HashSet<String> userID = new LinkedHashSet();

		for (int i = 0; i < dealerinfo.size(); i++) {

			String id = (String) dealerinfo.get(i).get("Dealer_User_id");
			String flagExecutionControl = (String) dealerinfo.get(i).get("Execution Control");

			if (!flagExecutionControl.equalsIgnoreCase("NO")) {
			userID.add(id);
			}
		}
		// System.out.println("###########"+userID);
		return userID;

	}

/*	public static String getflagAddtionalTestCase() throws Throwable {

		String flagAddtionalTestcases = "No";

		for (int i = 0; i < dealerinfo.size(); i++) {

			String racfID = (String) dealerinfo.get(i).get("Dealer_User_id");

			if (racfID.equalsIgnoreCase(strUserRACFID.toString().trim())) {

				flagAddtionalTestcases = (String) dealerinfo.get(i).get("Addtional Testcases");
				break;
			}
		}
		return flagAddtionalTestcases.trim();

	}*/

	public static void getDealersInfoFlag() throws Throwable {
		flagDealerPrinciple = "No";
		flagAddtionalTestcases = "No";
		flagDealerType= "NA";
	
		for (int i = 0; i < dealerinfo.size(); i++) {
			String racfID = (String) dealerinfo.get(i).get("Dealer_User_id");
			if (racfID.equalsIgnoreCase(strUserRACFID.toString().trim())) {
				flagDealerPrinciple = (String) dealerinfo.get(i).get("Dealer Principle").toString().trim();
				flagAddtionalTestcases = (String) dealerinfo.get(i).get("Addtional Testcases").toString().trim();
				flagDealerType = (String) dealerinfo.get(i).get("Dealer Type (Main/Sub)").toString().trim();
				LogFactory.info("DealerPrinciple : " + flagDealerPrinciple + "AddtionalTestcases : " +flagAddtionalTestcases  + "DealerType : " +flagDealerType );
				break;
			}
		}
	}

	@SuppressWarnings("rawtypes")
	public static void readWCMContentData() throws Exception {
		try {
			ArrayList<String> ListAttributeName = new ArrayList<String>();
			FileInputStream inputStream = new FileInputStream(new File(strDataPath));
			book = new XSSFWorkbook(inputStream);
			
			dataDealerInfoSheet = book.getSheet("Dealer Info");
			dataWCMContentSheet = book.getSheet("WCM Content");
			dataAdditionalTCSheet = book.getSheet("Additional Testcases");
			dataTranslationsSheet = book.getSheet("Translations");
			
			int Row_count = dataWCMContentSheet.getLastRowNum();
			int colon_count = dataWCMContentSheet.getRow(0).getLastCellNum();
			String AttributeName = "";
			String AttributeValue = "";
			String Dealer_User_value = "";
			String Coutryvalue = "";
			String Dealer_ProductType = "";

			for (int j = 0; j < colon_count; j++) {
				AttributeName = dataWCMContentSheet.getRow(0).getCell(j).getStringCellValue().trim();
				ListAttributeName.add(AttributeName);
			}
			dealerinfo = readDealerInfoData();
			wcmExceldata = new ArrayList<LinkedHashMap>();
			for (int k = 0; k < dealerinfo.size(); k++)
			{
				Dealer_User_value = (String) dealerinfo.get(k).get("Dealer_User_id").toString().trim();
				Coutryvalue = (String) dealerinfo.get(k).get("Dealer_Country").toString().trim();
				Dealer_ProductType = (String) dealerinfo.get(k).get("Dealer_ProductType").toString().trim();
				for (int i = 1; i <= Row_count; i++) {
					LinkedHashMap<String, String> RowValuedata = new LinkedHashMap<String, String>();
					RowValuedata.put("Dealer_User_id", Dealer_User_value);
					RowValuedata.put("Dealer_Country", Coutryvalue);
					RowValuedata.put("Dealer_ProductType", Dealer_ProductType);
					for (int j = 0; j < colon_count; j++) {
						AttributeName = ListAttributeName.get(j);
						AttributeValue = dataWCMContentSheet.getRow(i).getCell(j).getStringCellValue().trim();
						RowValuedata.put(AttributeName, AttributeValue);
					}
					wcmExceldata.add(RowValuedata);
				}
			}

		} catch (Exception e) {
			System.out.println(e.getMessage());

		}
	}

	@SuppressWarnings("unused")
	public static void getUserWCMContent() throws Throwable {

		// String strUserRACFID = "XCI8711";

		userWCMData = new ArrayList<LinkedHashMap>();

		for (int i = 0; i < wcmExceldata.size(); i++) {

			String racfID = (String) wcmExceldata.get(i).get("Dealer_User_id");

			if (racfID.equalsIgnoreCase(strUserRACFID.toString().trim())) {

				userWCMData.add(wcmExceldata.get(i));
			}
		}
		// System.out.println(userWCMData);

	}

	public static void writeExcel() {

		try {

			DateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy_HH-mm-ss");
			Date date = new Date();
			String filename = strDataOutputPath + dateFormat.format(date) + ".xlsx";

			ExcelFactory.writeHeader(filename, BaseClass.headerList);

			ExcelFactory.writeRow(filename, BaseClass.finalResultforExcel);
		} catch (InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();

		}
	}
	
	@SuppressWarnings("rawtypes")
	 public static List<LinkedHashMap> getUserWcmDetailsAfterFilteringCountryAndProduct(String ContentType, String IndexPage, String ChildIndexPage,String GrandChildIndexPage) {
	  LinkedHashMap rowdata = null;
	  List<LinkedHashMap> rowDataList = new ArrayList<LinkedHashMap>();
	  for (int i = 0; i < userWCMData.size(); i++) {
	   rowdata = userWCMData.get(i);
	   String id_value = (String) rowdata.get("Dealer_User_id");
	   if (id_value.equalsIgnoreCase(strUserRACFID)) {
	    String value = (String) rowdata.get("ContentType");
	    String dealerCountry = (String) rowdata.get("Dealer_Country");
	    String delarProdType = (String) rowdata.get("Dealer_ProductType");
	    String wcmcountry = (String) rowdata.get("Country");
	    String wcmdProdType = (String) rowdata.get("ProductType");
	    if (GenericFactory.userAndWCMCountryComparison(dealerCountry, wcmcountry)
	      && GenericFactory.userAndWCMProductTypeComparison(delarProdType, wcmdProdType)
	      && value.equalsIgnoreCase(ContentType) || value.equalsIgnoreCase(IndexPage) || value.equals(ChildIndexPage) || value.equalsIgnoreCase(GrandChildIndexPage)) {
	     rowDataList.add(rowdata);
	     System.out.println(rowdata.get("Country"));
	    }
	   }
	  }
	  return rowDataList;
	 }
	
	@SuppressWarnings("rawtypes")
	 public static List<LinkedHashMap> getUserWcmDetailsAfterFilteringCountryAndProduct(String ContentType) {
	  LinkedHashMap rowdata = null;
	  List<LinkedHashMap> rowDataList = new ArrayList<LinkedHashMap>();
	  
	  for (int i = 0; i < userWCMData.size(); i++) {
	   rowdata = userWCMData.get(i);
	   String id_value = (String) rowdata.get("Dealer_User_id");
	   if (id_value.equalsIgnoreCase(strUserRACFID)) {
	    String value = (String) rowdata.get("ContentType");
	    String dealerCountry = (String) rowdata.get("Dealer_Country");
	    String delarProdType = (String) rowdata.get("Dealer_ProductType");
	    String wcmcountry = (String) rowdata.get("Country");
	    String wcmdProdType = (String) rowdata.get("ProductType");
	    
	    if (GenericFactory.userAndWCMCountryComparison(dealerCountry, wcmcountry)
	      && GenericFactory.userAndWCMProductTypeComparison(delarProdType, wcmdProdType)
	      && value.equalsIgnoreCase(ContentType)) {
	    	
	     rowDataList.add(rowdata);
	     
	    }
	   }
	  }
	  return rowDataList;
	 }
	
	

}
