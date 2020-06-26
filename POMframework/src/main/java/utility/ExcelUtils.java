package utility;
import java.awt.Color;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;

public class ExcelUtils {

	private static XSSFSheet ExcelWSheet;
	private static XSSFWorkbook ExcelWBook;
	private static XSSFCell Cell;
	private static XSSFRow Row;
	private static MissingCellPolicy xRow;

	//This method is to set the File path and to open the Excel file, Pass Excel Path and Sheetname as Arguments to this method
	public static void setExcelFile(String Path,String SheetName) throws Exception {
		try {			
			// Open the Excel file			
			File src= new File("./testData/TestData.xlsx");
			FileInputStream ExcelFile = new FileInputStream(src);					
			// Access the required test data sheet
			ExcelWBook = new XSSFWorkbook(ExcelFile);			
			ExcelWSheet = ExcelWBook.getSheet(SheetName);			
			Log.info("Excel sheet opened");

		} catch (Exception e){
			throw (e);
		}
	}
	//This method is to read the test data from the Excel cell, in this we are passing parameters as Row num and Col num
	public static String getCellData(int RowNum, int ColNum) throws Exception{
		try{
			Cell = ExcelWSheet.getRow(RowNum).getCell(ColNum);
			String CellData = Cell.getStringCellValue();
			return CellData;
		}catch (Exception e){
			return"";
		}
	}
	//This method is to write in the Excel cell, Row num and Col num are the parameters
	
	
	@SuppressWarnings("static-access")
	public static void setCellData(String Result,  int RowNum, int ColNum) throws Exception   
	{
		XSSFFont fontBold = ExcelUtils.ExcelWBook.createFont()  ;    //oWB.createFont();
		XSSFFont fontWhite = ExcelUtils.ExcelWBook.createFont();
		XSSFColor WhiteColor = new XSSFColor(Color.WHITE);
		XSSFFont fontGreen = ExcelUtils.ExcelWBook.createFont();
		XSSFColor GreenColor = new XSSFColor(Color.GREEN);
		XSSFCellStyle styleGreen = ExcelUtils.ExcelWBook.createCellStyle();
		XSSFFont fontRed = ExcelUtils.ExcelWBook.createFont();
		XSSFColor RedColor = new XSSFColor(Color.RED);
		XSSFCellStyle styleRed = ExcelUtils.ExcelWBook.createCellStyle();
		XSSFColor YellowColor = new XSSFColor(Color.yellow);
		XSSFCellStyle styleYellow = ExcelUtils.ExcelWBook.createCellStyle();
		
		try{
			Row  = ExcelWSheet.getRow(RowNum);
			//Cell = Row.getCell(ColNum, Row.RETURN_BLANK_AS_NULL);
			Cell = Row.getCell(ColNum, xRow.RETURN_BLANK_AS_NULL);
			//color the Result cell
			if(Result.trim().compareToIgnoreCase("Pass")==0)
			{
				fontGreen.setColor(new XSSFColor(new java.awt.Color(0,100,0)));
				styleGreen.setFont(fontGreen);
				styleGreen.setFillForegroundColor(GreenColor);							
				styleGreen.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				Cell.setCellStyle(styleGreen);
			}
			else if(Result.trim().compareToIgnoreCase("Fail")==0)
			{
				fontRed.setColor(RedColor);
				styleRed.setFont(fontRed);
				styleRed.setFillForegroundColor(RedColor);			
				styleRed.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				Cell.setCellStyle(styleRed);
			}
			//give result to the cell
			if (Cell == null) {
				Cell = Row.createCell(ColNum);
				Cell.setCellValue(Result);				
			} else {
				Cell.setCellValue(Result);				
			}
			// Constant variables Test Data path and Test Data file name
			File src= new File("./testData/TestData.xlsx");
			FileOutputStream fileOut = new FileOutputStream(src);			
			ExcelWBook.write(fileOut);
			fileOut.flush();
			fileOut.close();
		}catch(Exception e){
			throw (e);
		}
	}
	
	public static int getRowContains(String sTestCaseName, int colNum) throws Exception{
		int i;
		try {
			int rowCount = ExcelUtils.getRowUsed();
			for ( i=0 ; i<rowCount; i++){
				if  (ExcelUtils.getCellData(i,colNum).equalsIgnoreCase(sTestCaseName)){
					break;
				}
			}
			return i;
		}catch (Exception e){
			Log.error("Class ExcelUtil | Method getRowContains | Exception desc : " + e.getMessage());
			throw(e);
		}
	}

	public static int getRowUsed() throws Exception {
		try{
			int RowCount = ExcelWSheet.getLastRowNum();
			Log.info("Total number of Row used return as < " + RowCount + " >.");		
			return RowCount;
		}catch (Exception e){
			Log.error("Class ExcelUtil | Method getRowUsed | Exception desc : "+e.getMessage());
			System.out.println(e.getMessage());
			throw (e);
		}

	}
}