package com.prem;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.stream.StreamSupport;

import org.apache.commons.collections4.map.HashedMap;
import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.opencsv.CSVReader;

public class CsvToExcel {

	/*
	 * 
	 * public static String ConvertCSVToXLS(String file) throws IOException {
	 * 
	 * if (file.indexOf(".csv") < 0) return
	 * "Error converting file: .csv file not given.";
	 * 
	 * String name = FileManager.getFileNameFromPath(file, false);
	 * ArrayList<ArrayList<String>> arList = new ArrayList<ArrayList<String>>();
	 * ArrayList<String> al = null;
	 * 
	 * String thisLine; DataInputStream myInput = new DataInputStream(new
	 * FileInputStream(file));
	 * 
	 * while ((thisLine = myInput.readLine()) != null) { al = new
	 * ArrayList<String>(); String strar[] = thisLine.split(",");
	 * 
	 * for (int j = 0; j < strar.length; j++) { // My Attempt (BELOW) String
	 * edit = strar[j].replace('\n', ' '); al.add(edit); }
	 * 
	 * arList.add(al); System.out.println(); }
	 * 
	 * try { HSSFWorkbook hwb = new HSSFWorkbook(); HSSFSheet sheet =
	 * hwb.createSheet("new sheet");
	 * 
	 * for (int k = 0; k < arList.size(); k++) { ArrayList<String> ardata =
	 * (ArrayList<String>) arList.get(k); HSSFRow row = sheet.createRow((short)
	 * 0 + k);
	 * 
	 * for (int p = 0; p < ardata.size(); p++) {
	 * System.out.print(ardata.get(p)); HSSFCell cell = row.createCell((short)
	 * p); cell.setCellValue(ardata.get(p).toString()); } }
	 * 
	 * FileOutputStream fileOut = new FileOutputStream(
	 * FileManager.getCleanPath() + "/converted files/" + name + ".xls");
	 * hwb.write(fileOut); fileOut.close();
	 * 
	 * System.out.println(name + ".xls has been generated"); } catch (Exception
	 * ex) { }
	 * 
	 * return ""; }
	 */

	public static final char FILE_DELIMITER = ',';
	public static final String FILE_EXTN = ".xlsx";
	public static final String FILE_NAME = "EXCEL_DATA";

	private static Logger logger = Logger.getLogger(CsvToExcel.class);

	public static String convertCsvToXls(String xlsFileLocation, String csvFilePath) {
		SXSSFSheet sheet = null;
		CSVReader reader = null;
		Workbook workBook = null;
		String generatedXlsFilePath = "";
		FileOutputStream fileOutputStream = null;

		try {

			/****
			 * Get the CSVReader Instance & Specify The Delimiter To Be Used
			 ****/
			String[] nextLine;
			reader = new CSVReader(new FileReader(csvFilePath), FILE_DELIMITER);

			workBook = new SXSSFWorkbook();
			sheet = (SXSSFSheet) workBook.createSheet("Sheet");

			int rowNum = 0;
			logger.info("Creating New .Xls File From The Already Generated .Csv File");
			while ((nextLine = reader.readNext()) != null) {
				Row currentRow = sheet.createRow(rowNum++);
				for (int i = 0; i < nextLine.length; i++) {
					// System.out.println("Row is::" + i + " Value is:: " +
					// nextLine[i]);

					/*
					 * if(NumberUtils.isDigits(nextLine[i])) {
					 * currentRow.createCell(i).setCellValue(Integer.parseInt(
					 * nextLine[i]));
					 * 
					 * } else if (NumberUtils.isNumber(nextLine[i])) {
					 * currentRow.createCell(i).setCellValue(Double.parseDouble(
					 * nextLine[i])); } else {
					 * currentRow.createCell(i).setCellValue(nextLine[i]); }
					 */
					currentRow.createCell(i).setCellValue(nextLine[i]);
				}
			}

			generatedXlsFilePath = xlsFileLocation + FILE_NAME + FILE_EXTN;
			logger.info("The File Is Generated At The Following Location?= " + generatedXlsFilePath);

			fileOutputStream = new FileOutputStream(generatedXlsFilePath.trim());
			workBook.write(fileOutputStream);
		} catch (Exception exObj) {
			logger.error("Exception In convertCsvToXls() Method?=  " + exObj);
		} finally {
			try {

				workBook.close();
				fileOutputStream.close();
				reader.close();
			} catch (IOException ioExObj) {
				logger.error("Exception While Closing I/O Objects In convertCsvToXls() Method?=  " + ioExObj);
			}
		}

		return generatedXlsFilePath;
	}

	public static void readExcel(String filePath) {
		try {
			// File file = new File("C:\\mysheets\\student.xlsx");
			File file = new File(filePath);
			// creating a new file instance
			FileInputStream fis = new FileInputStream(file);
			// obtaining bytes from the file
			// creating Workbook instance that refers to .xlsx file
			XSSFWorkbook wb = new XSSFWorkbook(fis);
			XSSFSheet sheet = wb.getSheetAt(0);
			// creating a Sheet object to retrieve object
			Iterator<Row> itr = sheet.iterator();
			// iterating over excel file
			System.out.println("The given file is");
			while (itr.hasNext()) {
				Row row = itr.next();
				Iterator<Cell> cellIterator = row.cellIterator();
				// iterating over each column
				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					switch (cell.getCellType()) {
					case Cell.CELL_TYPE_STRING:
						// field that represents string cell type
						System.out.print(cell.getStringCellValue() + "\t\t\t");
						break;
					case Cell.CELL_TYPE_NUMERIC:
						// field that represents number cell type
						System.out.print(cell.getNumericCellValue() + "\t\t\t");
						break;
					default:
					}
				}
				System.out.println("");
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static Map<String,Set<Object>> mapUniqueExcelColumnData(String filePath) {
		String value,keyValue = null; // variable for storing the cell value
		Workbook wbook = null; // initialize Workbook null
		Map<String, Set<Object>> map = new HashMap<String, Set<Object>>();
		try {
			
			File file = new File(filePath);
			FileInputStream fis = new FileInputStream(file);
			wbook = new XSSFWorkbook(fis);
			Sheet sheet = wbook.getSheetAt(0);
			Iterator<Row> itr = sheet.iterator();

			Row nextRow = itr.next();
			int rowCount = sheet.getLastRowNum();
			int columnCount = nextRow.getLastCellNum();

			System.out.println("total row is ::" + rowCount + " toatl column is::" + columnCount);
			
			for(int i=0;i<columnCount;i++){
				Set<Object> list=new HashSet<Object>();
				for(int j=1;j<=rowCount;j++){
					Row row2 = sheet.getRow(j);
				    Cell cell2 = row2.getCell(i);
				    value = cell2.getStringCellValue();
				    list.add(value);
				}
				Row rowKey = sheet.getRow(0);
			    Cell cellKey = rowKey.getCell(i);
			    keyValue = cellKey.getStringCellValue();
			    map.put(keyValue, list);
			}
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e1) {
			e1.printStackTrace();
		}

		return map;
		// returns the map which contain data in column wise.

	}
	
	public static Map<String,List<Object>> mapAllExcelColumnData(String filePath) {
		String value,keyValue = null; // variable for storing the cell value
		Workbook wbook = null; // initialize Workbook null
		Map<String, List<Object>> map = new HashMap<String, List<Object>>();
		try {
			
			File file = new File(filePath);
			FileInputStream fis = new FileInputStream(file);
			wbook = new XSSFWorkbook(fis);
			Sheet sheet = wbook.getSheetAt(0);
			Iterator<Row> itr = sheet.iterator();

			Row nextRow = itr.next();
			int rowCount = sheet.getLastRowNum();
			int columnCount = nextRow.getLastCellNum();

			System.out.println("total row is ::" + rowCount + " toatl column is::" + columnCount);
			
			for(int i=0;i<columnCount;i++){
				List<Object> list=new ArrayList<Object>();
				for(int j=1;j<=rowCount;j++){
					Row row2 = sheet.getRow(j);
				    Cell cell2 = row2.getCell(i);
				    value = cell2.getStringCellValue();
				    list.add(value);
				}
				Row rowKey = sheet.getRow(0);
			    Cell cellKey = rowKey.getCell(i);
			    keyValue = cellKey.getStringCellValue();
			    map.put(keyValue, list);
			}
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e1) {
			e1.printStackTrace();
		}

		return map;
		// returns the map which contain data in column wise.

	}

	
	

	public static void main(String[] args) {

		String xlsLoc = "config/", csvLoc = "config/sample.csv", fileLoc = "";
		fileLoc = CsvToExcel.convertCsvToXls(xlsLoc, csvLoc);
		 readExcel(fileLoc);

		Map<String,Set<Object>> columnMap = mapUniqueExcelColumnData(fileLoc);
		System.out.println(columnMap);
		System.out.println("\n File Location Is?= " + fileLoc);
		System.out.println("Vehicle no is::"+columnMap.get("VEHICLE_NO"));
		System.out.println("Vehicle count is::"+columnMap.get("VEHICLE_NO").size());
		
		
		
	}

}
