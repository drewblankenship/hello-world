package com.nationwide.nf.pressys.excel;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.sql.Timestamp;
import java.util.Iterator;
import java.util.Properties;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {
	public static String[][] data;
	public static String[][] dataFound;
	public static String[][] multipleSearchData;
	public static String[][] information;
	public static String[] category, roles, projects, lobs, corrTypes;
	public static String value = "";
	public static int rowNum = 0;
	public static int cellNum = 0;
	public static String tempOption = "";
	static String propertyFileLocation = "C:\\Users\\blankd2\\documents\\On_Call_Run\\On_Call_Run\\ExcelFiles.properties";
	public static Properties props;
	public void read() throws IOException{
		props = new Properties();
		props.load(new FileReader(propertyFileLocation));
		FileInputStream file = new FileInputStream(new File(props.getProperty("excelFile")));
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheetAt(1);
		Iterator<Row> rowIterator = sheet.iterator();
		rowNum = sheet.getPhysicalNumberOfRows();
		cellNum = 9;
		System.out.println("Num Of Rows: " + rowNum);
		System.out.println("Num Of Cells: " + cellNum);

		data = new String[rowNum][cellNum];
		dataFound = new String[rowNum][cellNum];
		category = new String[cellNum];
		information = new String[rowNum][cellNum];
		int tempRow = 0;
		int tempCell = 0;
		int tempInfoRow = 0;
		while(rowIterator.hasNext()){
			Row row = rowIterator.next();
			Iterator<Cell> cellIterator = row.cellIterator();
			//System.out.println("Row: " + tempRow);
			tempCell = 0;
			while(/*cellIterator.hasNext()||*/tempCell < cellNum){
				Cell cell = row.getCell(tempCell);
				if(cell != null){
						switch(cell.getCellType()){
						case Cell.CELL_TYPE_NUMERIC:
							value = String.valueOf((int) cell.getNumericCellValue());
							//System.out.println("Value: " + value);
							break;
						case Cell.CELL_TYPE_STRING:
							value = cell.getStringCellValue();
							//System.out.println("Value: " + value);
							break;
						case Cell.CELL_TYPE_BLANK:
							value = "";
						}
				}
				if(tempCell == (cellNum - 1)){
					value = String.valueOf(tempRow);
				}
				if(tempRow > 0){
					information[tempRow][tempCell] = value;
					
				}
				data[tempRow][tempCell] = value;
				tempCell++;
			}
			tempRow++;
		}
		String[][] tempInfo = new String[rowNum-1][cellNum];
		for(int i = 0; i < tempInfo.length; i++){
			tempInfo[i] = information[i+1];
		}
		information = new String[rowNum-1][cellNum];
		information = tempInfo;
		/*for(int i = 0; i < information.length; i++){
			information[i] = tempInfo[i];
		}*/
		for(int i = 0; i <data[0].length; i++){
			category[i] = data[0][i];
		}
		workbook.close();
		file.close();
	}
	public void search(int searchCategory, String searchingFor) throws IOException{
		dataFound = new String[rowNum][cellNum];
		int foundCount = 0;
		for(int i = 0; i < data.length; i++){
			if(data[i][searchCategory].contains(searchingFor)){
				for(int j = 0; j < data[i].length; j++){
					dataFound[foundCount][j] = data[i][j];
					//System.out.println("Data Found: " + dataFound[foundCount][j]);
				}
				foundCount++;
			}
		}
	}
	public void multipleSearch(int[] searchCategory, String[] searchingFor) throws IOException{
		int foundCount = 0;
		int tempCount = 0;
		int tempCategory = 0;
		String[][] oldData = new String[rowNum][cellNum];
		String[][] tempData = new String[rowNum][cellNum];
		for(int x = 0; x < searchCategory.length; x++){
			int searchCriteria = searchCategory[x];
			String searched = searchingFor[x];
			foundCount = 0;
			tempCount = 0;
			if(x == 0){
				tempData = new String[data.length][data[0].length];
				for(int i = 0; i < data.length; i++){
					if(data[i][searchCriteria].contains(searched)){
						foundCount++;
					}
				}
				tempData = data;
				dataFound = new String[foundCount][cellNum];
				oldData = new String[foundCount][cellNum];
				for(int i = 0; i <data.length; i++){
					if(data[i][searchCriteria].contains(searched)){
						for(int j = 0; j < data[i].length; j++){
							dataFound[tempCount][j] = data[i][j];
						}
						tempCount++;
					}
				}
				oldData = dataFound;
				tempCategory = searchCategory[0];
			}else{
				if(tempCategory == searchCategory[x]){
					for(int i = 0; i < tempData.length; i++){
						if(tempData[i][searchCriteria].contains(searched)){
							foundCount++;
						}
					}
					multipleSearchData = new String[foundCount][cellNum];
					for(int i = 0; i < tempData.length; i++){
						if(tempData[i][searchCriteria].contains(searched)){
							for(int j = 0; j < tempData[i].length; j++){
								multipleSearchData[tempCount][j] = tempData[i][j];
							}
							tempCount++;
						}
					}
					dataFound = new String[oldData.length + multipleSearchData.length][cellNum];
					for(int i = 0; i < oldData.length; i++){
						for(int j = 0; j < cellNum; j++){
							dataFound[i][j] = oldData[i][j];
						}
					}
					for(int i = 0; i<multipleSearchData.length; i++){
						for(int j = 0; j < cellNum; j++){
							dataFound[oldData.length+i][j] = multipleSearchData[i][j];
						}
					}
				}else{
					for(int i = 0; i < dataFound.length; i++){
						if(dataFound[i][searchCriteria].contains(searched)){
							foundCount++;
						}
					}
					tempData = new String[dataFound.length][dataFound[0].length];
					tempData = dataFound;
					multipleSearchData = new String[foundCount][cellNum];
					for(int i = 0; i < dataFound.length; i++){
						if(dataFound[i][searchCriteria].contains(searched)){
							for(int j = 0; j < dataFound[i].length; j++){
								multipleSearchData[tempCount][j] = dataFound[i][j];
							}
							tempCount++;
						}
					}
					dataFound = new String[foundCount][cellNum];
					oldData = new String[foundCount][cellNum];
					for(int i = 0; i < foundCount; i++){
						for(int j = 0; j < cellNum; j++){
							dataFound[i][j] = multipleSearchData[i][j];
						}
					}
					oldData = dataFound;
					tempCategory = searchCategory[x];
				}
			}	
		}
	}
	public String[][] getData(){
		return data;
	}
	public String[][] getDataFound(){
		return dataFound;
	}
	public String[][] getMultipleSearchData(){
		return multipleSearchData;
	}
	public String[][] getInformation(){
		return information;
	}
	public String[] getCategory(){
		return category;
	}
	public int getRowNum(){
		return rowNum;
	}
	public int getCellNum(){
		return cellNum;
	}
	public String[] getRoles(){
		return roles;
	}
	public String[] getProjects(){
		return projects;
	}
	public String[] getLOBs(){
		return lobs;
	}
	public String[] getCorrTypes(){
		return corrTypes;
	}
	public String getFileLocation(){
		return props.getProperty("excelFile");
	}
}
