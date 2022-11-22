package common;
import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;

import entity.ExcelBook;
import entity.ExcelSheet;
import entity.RowObj;
public class ReadExcel {
	
	public static ExcelBook read (File file) {
		
		ExcelBook excelBook = new ExcelBook();
		
		try {
			InputStream fis = new FileInputStream(file.getAbsolutePath());
			
			BufferedInputStream bis = new BufferedInputStream(fis);
	        if(POIFSFileSystem.hasPOIFSHeader(bis)) {
	        	excelBook = readForXLS(file);
	        }
			/*if(POIXMLDocument.hasOOXMLHeader(bis)) {
			    System.out.println("2007及以上");
			}*/
	        
			
		} catch(Exception e) {
			System.out.println(e.toString());
		} finally {
			
		}
		return excelBook;
	}
	
	/**
	 * 读取Excel文件，以java对象存储。
	 * 最大支持80列数据的读取，每行数据存储在RowObj对象中
	 * @param file 文件对象
	 * @return Excel按Sheet分所有的内容
	 * @throws IOException
	 */
	private static ExcelBook readForXLS(File file) throws IOException {
		HSSFWorkbook workbook = null;
		int numberOfSheet = 0;
		int lastCellNumber = 80;
		ExcelBook excelBook = new ExcelBook();
		try {
			// 获取Excel对象
			workbook = new HSSFWorkbook(new FileInputStream(file.getAbsolutePath()));
			
			Map<String, ExcelSheet> excelSheetContent = new HashMap<String, ExcelSheet>();
			
			// 获取Excel中Sheet的数量
			numberOfSheet = workbook.getNumberOfSheets();
			
			// 遍历Sheet，获取数据封装到对象中
			for (int sheetNmu = 0; sheetNmu < numberOfSheet; sheetNmu++) {
				ExcelSheet excelSheet = new ExcelSheet();
				List<RowObj> rowList = new ArrayList<RowObj>();
				// 行遍历
				HSSFSheet sheet = workbook.getSheetAt(sheetNmu);
				for(int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum++) {
					RowObj rowObj = new RowObj();
					
					HSSFRow row = sheet.getRow(rowNum);
					if (row != null) {
						rowObj.setColCount(row.getLastCellNum());
					} else {
						continue;
					}
					// 列遍历
					for (int cellNum = 0; cellNum <= lastCellNumber; cellNum++) {
						// 通过属性名循环设定属性值
						Field field = rowObj.getClass().getDeclaredField("column" + cellNum);
						if (field != null && row.getCell(cellNum) != null) {
							field.setAccessible(true);
							row.getCell(cellNum).setCellType(Cell.CELL_TYPE_STRING);
							field.set(rowObj, row.getCell(cellNum).getStringCellValue() != null ? row.getCell(cellNum).getStringCellValue() : "");
							rowObj.setRowNumber(rowNum);
						} else {
							continue;
						}
						
					}
					rowList.add(rowObj);
				}
				// sheet名
				excelSheet.setSheetName(sheet.getSheetName());
				// sheet内容
				excelSheet.setRowObj(rowList);
				
				// book内容
				excelSheetContent.put(sheet.getSheetName(), excelSheet);
				excelBook.setExcelSheets(excelSheetContent);
				excelBook.setBookName(file.getName());
			}
		} catch(Exception e) {
			System.out.println(e.toString());
		} finally {
			workbook.close();
		}
		return excelBook;
	}
}
