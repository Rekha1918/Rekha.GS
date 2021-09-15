package test.java.excelExportAndFileIO;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;

public class WriteExcelFile extends ReadExcelFile {
//	Workbook iQuantiWorkbook = null;
//	File file;
	protected static FileOutputStream foutput;
	protected static String filePath2="";
	protected static String fileName2="";
	protected static int masterExcelRowNum;
	protected static int masterExcelColumnNum;
	
	Sheet iQuantiSheet2;
//	protected static FileOutputStream fileOutput;
	ReadExcelFile r=new ReadExcelFile();
	

//	public Sheet openZerothExcel() throws IOException {
//		return iQuantiSheet = r.readExcel(System.getProperty("user.dir")+"\\","TestCase.xlsx" , "0");
//	}

//	public Sheet openFirstExcel() throws IOException {
//		return iQuantiSheet = r.readExcel(System.getProperty("user.dir")+"\\","TestCase.xlsx" , "1");
//	}
//	
//	public Sheet openSecondExcel() throws IOException {
//		return iQuantiSheet = r.readExcel(System.getProperty("user.dir")+"\\","TestCase.xlsx" , "2");
//	}
//	
//	public Sheet openThirdExcel() throws IOException {
//		return iQuantiSheet = r.readExcel(System.getProperty("user.dir")+"\\","TestCase.xlsx" , "3");
//	}
//	
//	public Sheet openFourthExcel() throws IOException {
//		return iQuantiSheet = r.readExcel(System.getProperty("user.dir")+"\\","TestCase.xlsx" , "4");
//	}
	
	public void openMasterExcel() throws IOException
	{
		Sheet iQuantiSheet2 = r.readExcel(System.getProperty("user.dir")+"\\","MasterTestSuite.xlsx" , "0");
		iQuantiSheet2.getRow(1).getCell(1).setCellType(CellType.STRING);
		iQuantiSheet2.getRow(1).getCell(2).setCellType(CellType.STRING);
		masterExcelRowNum=Integer.parseInt(iQuantiSheet2.getRow(1).getCell(1).getStringCellValue());
		masterExcelColumnNum=Integer.parseInt(iQuantiSheet2.getRow(1).getCell(2).getStringCellValue());
		
		filePath2=iQuantiSheet2.getRow(masterExcelRowNum).getCell(masterExcelColumnNum).getStringCellValue();
		fileName2=iQuantiSheet2.getRow(masterExcelRowNum).getCell(masterExcelColumnNum+1).getStringCellValue();
	}
	
	public Sheet openZerothExcel() throws IOException {
		return iQuantiSheet = r.readExcel((System.getProperty("user.dir")+"\\"+filePath2), fileName2 , "0");
	}
	
	public Sheet openFirstExcel() throws IOException {
		return iQuantiSheet = r.readExcel((System.getProperty("user.dir")+"\\"+filePath2), fileName2 , "1");
	}
	
	public Sheet openSecondExcel() throws IOException {
		return iQuantiSheet = r.readExcel((System.getProperty("user.dir")+"\\"+filePath2), fileName2 , "2");
	}
	
	public Sheet openThirdExcel() throws IOException {
		return iQuantiSheet = r.readExcel((System.getProperty("user.dir")+"\\"+filePath2), fileName2 , "3");
	}
	
	public Sheet openFourthExcel() throws IOException {
		return iQuantiSheet = r.readExcel((System.getProperty("user.dir")+"\\"+filePath2), fileName2 , "4");
	}
	
	public Sheet openFifthExcel() throws IOException {
		return iQuantiSheet = r.readExcel((System.getProperty("user.dir")+"\\"+filePath2), fileName2 , "5");
	}
	
	
	public void writeZerothExcelStatus(int rowNum, String status) throws IOException {
		foutput = new FileOutputStream(file);
		iQuantiSheet.getRow(rowNum).getCell(10).setCellValue(status);
		iQuantiWorkbook.write(foutput);
	}
	
	public void writeActualGetText(int rowNum, String actualGetText) throws IOException {
		foutput = new FileOutputStream(file);
		iQuantiSheet.getRow(rowNum).getCell(9).setCellValue(actualGetText);
		iQuantiWorkbook.write(foutput);
	}
	
	public void writeActualGetText2(int trackFourthSheetRow, int actualGetTextColumnCount, String actualGetText2, String status) throws IOException {
		foutput = new FileOutputStream(file);
		iQuantiSheet.getRow(trackFourthSheetRow).getCell(actualGetTextColumnCount).setCellValue(actualGetText2);
		iQuantiSheet.getRow(trackFourthSheetRow).getCell(actualGetTextColumnCount+1).setCellValue(status);
		iQuantiWorkbook.write(foutput);
	}
	
	public void writeSelectedDropdownOption(int rowNum, String selectedOption) throws IOException {
		foutput = new FileOutputStream(file);
		iQuantiSheet.getRow(rowNum).getCell(9).setCellValue(selectedOption);
		iQuantiWorkbook.write(foutput);
	}
	
	public void writeH1Tag(int sheetTwoRow, String actualH1Tag, String h1TagStatus) throws IOException {
		foutput = new FileOutputStream(file);
		iQuantiSheet.getRow(sheetTwoRow).getCell(2).setCellValue(actualH1Tag);
		iQuantiSheet.getRow(sheetTwoRow).getCell(3).setCellValue(h1TagStatus);
		iQuantiWorkbook.write(foutput);
	}
	
	public void writeMetaTag(int urlRowNum, String actualMetaTag, String metaTagStatus) throws IOException {
		foutput = new FileOutputStream(file);
		iQuantiSheet.getRow(urlRowNum).getCell(5).setCellValue(actualMetaTag);
		iQuantiSheet.getRow(urlRowNum).getCell(6).setCellValue(metaTagStatus);
		iQuantiWorkbook.write(foutput);
	}
	
	public void writePageTitle(int urlRowNum, String actualPageTitle, String pageTitleStatus) throws IOException {
		foutput = new FileOutputStream(file);
		iQuantiSheet.getRow(urlRowNum).getCell(8).setCellValue(actualPageTitle);
		iQuantiSheet.getRow(urlRowNum).getCell(9).setCellValue(pageTitleStatus);
		iQuantiWorkbook.write(foutput);
	}
	
	public void writeRedirectedUrl(int urlRowNum, String actualRedirect, String redirectStatus) throws IOException {
		foutput = new FileOutputStream(file);
		iQuantiSheet.getRow(urlRowNum).getCell(11).setCellValue(actualRedirect);
		iQuantiSheet.getRow(urlRowNum).getCell(12).setCellValue(redirectStatus);
		iQuantiWorkbook.write(foutput);
	}
	
	public void writeProductTags(int urlRowNum, int listCellNum2, String productTagsList, String productTagStatus) throws IOException {
		foutput = new FileOutputStream(file);
		iQuantiSheet.getRow(urlRowNum).getCell(listCellNum2+1).setCellValue(productTagsList);
		iQuantiSheet.getRow(urlRowNum).getCell(listCellNum2+2).setCellValue(productTagStatus);
		iQuantiWorkbook.write(foutput);
	}
	
	public void writeCompareMultipleList(int urlRowNum, int listCellNum, String allGetLinks, String compareMultipleListStatus) throws IOException {
		foutput = new FileOutputStream(file);
		iQuantiSheet.getRow(urlRowNum).getCell(listCellNum+1).setCellValue(allGetLinks);
		iQuantiSheet.getRow(urlRowNum).getCell(listCellNum+2).setCellValue(compareMultipleListStatus);
		iQuantiWorkbook.write(foutput);
	}
	
	public void closeExcel() throws IOException
	{
//		iQuantiWorkbook.write(foutput);
		foutput.close();
	}
//	public void closeExcel2() throws IOException
//	{
//		foutput = new FileOutputStream(file);
//		iQuantiSheet.getRow(1).getCell(10).setCellValue("PASS");
//		iQuantiWorkbook.write(foutput);
//		foutput.close();
//	}
}
