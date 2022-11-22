package common;

import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class WriteExcel {

	/**
	 * 将java对象写入Excel
	 * @param path 文件生成路径
	 * @param outFileName 文件名称
	 * @param title 输出文件内容的标题
	 * @param startRow 开始行
	 * @param resultList 输出对象List
	 * @throws IOException
	 * @throws IllegalArgumentException
	 * @throws IllegalAccessException
	 */
	public static void write(String path, String outFileName, String[] title, int startRow, List<?> resultList) throws IOException, IllegalArgumentException, IllegalAccessException {
		
		FileOutputStream fileOutputStream = new FileOutputStream(path + outFileName);
		int rowNumber = startRow + 1;
		Row rowContent = null;
		Cell cellContent = null;
		try {
			Workbook workbook = new HSSFWorkbook();
			
			Sheet sheet = workbook.createSheet("统计表");
			
			// 输出标题行
			Row row = sheet.createRow(startRow);
			for (int i = 0; i < title.length; i ++) {
				Cell cell = row.createCell(i);
				cell.setCellValue(title[i]);
			}
			
			// 输出统计内容 
			for(Object result : resultList) {
				Field[] property = result.getClass().getDeclaredFields();
				rowContent = sheet.createRow(rowNumber);
				for (int i = 0; i < property.length; i ++) {
					property[i].setAccessible(true);
					cellContent = rowContent.createCell(i);
					if (property[i].get(result) != null) {
						cellContent.setCellValue(property[i].get(result).toString());
					} else {
						cellContent.setCellValue("");
					}
				}
				rowNumber ++;
			}
			
			workbook.write(fileOutputStream);
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			fileOutputStream.close();
		}
	}
}
