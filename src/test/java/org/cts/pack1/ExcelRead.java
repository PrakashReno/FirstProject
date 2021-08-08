package org.cts.pack1;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRead {
	public static void main(String[] args) throws IOException {
		File file=new File("C:\\Users\\91880\\eclipse-workspace\\FaceBook\\TestData\\Emp.xlsx");
		FileInputStream stream=new FileInputStream(file);
		Workbook wb=new XSSFWorkbook(stream);
		Sheet sheet = wb.getSheet("Sheet1");
		for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
			Row row = sheet.getRow(i);
			for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
				Cell cell = row.getCell(j);
				int cellType = cell.getCellType();
				if (cellType==1) {
					String Value = cell.getStringCellValue();
					System.out.println(Value);
				}	else if (DateUtil.isCellDateFormatted(cell)) {
					Date dateCellValue = cell.getDateCellValue();
					SimpleDateFormat dateformat=new SimpleDateFormat("MM-DD-YY");
					String format = dateformat.format(dateCellValue);
					System.out.println(format);
				}
				else {
					double numericCellValue = cell.getNumericCellValue();
					long l=(long) numericCellValue;
					System.out.println(l);
				}
				
			}
			System.out.println();
		}
		
		
	}

}
