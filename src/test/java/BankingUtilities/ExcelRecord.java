package BankingUtilities;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRecord {

	private XSSFWorkbook workbook;
	private XSSFSheet sheet;
	private String filePath;
	private int rowNum;

	public ExcelRecord(String filePath, String sheetName) {
		this.filePath = filePath;

		try {
			FileInputStream fileInputStream = new FileInputStream(new File(filePath));
			workbook = new XSSFWorkbook(fileInputStream);
			sheet = workbook.getSheet(sheetName);
			fileInputStream.close();
		} catch (Exception e) {
			// TODO: handle exception
		}
	}
	
	public void writeData(String data) {
		if(sheet!=null) {
			Row row = sheet.createRow(rowNum++);
			Cell cell = row.createCell(0);
			cell.setCellValue(data);
		}
		
	}
	
	public void saveAndClose() {
		try {
		FileOutputStream fileOutputStream = new FileOutputStream(new File(filePath));
		workbook.write(fileOutputStream);
		fileOutputStream.close();
		workbook.close();
		}catch (Exception e) {
			// TODO: handle exception
		}
	}

}
