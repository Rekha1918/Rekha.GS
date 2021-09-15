package test.java.excelExportAndFileIO;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ReadExcelFile  {
	protected static File file;
	protected static Workbook iQuantiWorkbook = null;
	protected static FileInputStream inputStream;
	protected static Sheet  iQuantiSheet;
	
	public Sheet readExcel(String filePath,String fileName,String sheetNum) throws IOException{
	//Create a object of File class to open xlsx file
	file =	new File(filePath+"\\"+fileName);
	//Create an object of FileInputStream class to read excel file
	inputStream = new FileInputStream(file);
//	Workbook iQuantiWorkbook = null;
	//Find the file extension by spliting file name in substing and getting only extension name
	String fileExtensionName = fileName.substring(fileName.indexOf("."));
	//Check condition if the file is xlsx file
	if(fileExtensionName.equals(".xlsx")){
		//If it is xlsx file then create object of XSSFWorkbook class
		iQuantiWorkbook = new XSSFWorkbook(inputStream);
	}
	//Check condition if the file is xls file
	else if(fileExtensionName.equals(".xls")){
		//If it is xls file then create object of XSSFWorkbook class
		iQuantiWorkbook = new HSSFWorkbook(inputStream);
	}
	//Read sheet inside the workbook by its name
	Sheet  iQuantiSheet = iQuantiWorkbook.getSheetAt(Integer.parseInt(sheetNum));
	 return iQuantiSheet;	
	}
}
