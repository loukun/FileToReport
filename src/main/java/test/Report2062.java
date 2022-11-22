package test;

import java.io.File;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import common.FileTools;
import common.ReadExcel;
import common.StringUtils;
import common.WriteExcel;
import entity.ExcelBook;
import entity.ExcelSheet;
import entity.RowObj;
import result.Result2062;

public class Report2062 {
	
	public static void main(String[] args) throws NoSuchFieldException, SecurityException, IllegalArgumentException, IllegalAccessException, IOException {
		String path = "D:\\ExcelReadTest";
		List<File> outFileList = new ArrayList<File>();
		Map<String, List<Integer>> titleColumn = new HashMap<String, List<Integer>>();
		int startRow = 6;
		
		// 读取路径下所有的文件。
		FileTools.getAllFile(path, outFileList);
		
		List<ExcelBook> bookList = new ArrayList<ExcelBook>();
		ExcelBook excelBook = null;
		
		// 遍历Excel文件生成java对象
		for (File file : outFileList) {
			excelBook = ReadExcel.read(file);
			if (excelBook != null) {
				bookList.add(excelBook);
			}
		}
		
		// 遍历所有文件
		for (ExcelBook book : bookList) {
			Map<String, ExcelSheet> sheet = book.getExcelSheets();
			
			// 遍历文件中的Sheet页
			for (String key : sheet.keySet()) {
				List<Integer> list = new ArrayList<Integer>();
				RowObj rowObj = sheet.get(key).getRowObj().get(startRow);
				// 遍历标题行的所有列
				for (int j = 0; j <= rowObj.getColCount(); j ++) {
					Field f = rowObj.getClass().getDeclaredField("column" + j);
					f.setAccessible(true);
					if (f.get(rowObj) != null) {
						String title = f.get(rowObj).toString();
						if (title.contains("onclick")) {
//							System.out.println("sheet名：" + sheet.get(key).getSheetName() + " 列：" + j + " value:" +  f.get(rowObj));
							list.add(j);
						}
					}
				}
				// 按Sheet页统计onclick的列号
				if (list != null && list.size() > 0) {
					titleColumn.put(sheet.get(key).getSheetName(), list);
				}
			}
		}
		for (ExcelBook book : bookList) {
			Map<String, ExcelSheet> sheet = book.getExcelSheets();
			for (String key : sheet.keySet()) {
				String screenId = "";
				String screenName = "";
				
				for (int j = startRow; j < sheet.get(key).getRowObj().size(); j++) {
					if (sheet.get(key).getRowObj().get(j) != null) {
						if (sheet.get(key).getRowObj().get(j).getColumn2() != null && !"".equals(sheet.get(key).getRowObj().get(j).getColumn2()) && !screenId.equals(sheet.get(key).getRowObj().get(j).getColumn2())) {
							screenId = sheet.get(key).getRowObj().get(j).getColumn2();
							screenName = sheet.get(key).getRowObj().get(j).getColumn3();
						} else if (sheet.get(key).getRowObj().get(j).getColumn2() == null) {
							sheet.get(key).getRowObj().get(j).setColumn2(screenId);
							sheet.get(key).getRowObj().get(j).setColumn3(screenName);
						} else {
							sheet.get(key).getRowObj().get(j).setColumn2(screenId);
							sheet.get(key).getRowObj().get(j).setColumn3(screenName);
						}
					}
				}
			}
		}
		// 二次循环，提取统计列，生成统计结果
		List<Result2062> resultList = new ArrayList<Result2062>();
		for (ExcelBook book : bookList) {
			Map<String, ExcelSheet> sheet = book.getExcelSheets();
			for (String key : sheet.keySet()) {
				List<String> result = new ArrayList<String>();
				// 根据title列遍历所有行，取得对应值
				for (int j = startRow + 1; j < sheet.get(key).getRowObj().size(); j++) {
					RowObj row = sheet.get(key).getRowObj().get(j);
					
					List<Integer> list = titleColumn.get(sheet.get(key).getSheetName());
					if (list != null && list.size() > 0) {
						for (int k = 0; k < list.size(); k++) {
							Field f = row.getClass().getDeclaredField("column" + list.get(k));
							f.setAccessible(true);
							if (f.get(row) != null && !"".equals(f.get(row).toString())) {
								result.add(f.get(row).toString());
								//-----------------------------------------------------------------
								Result2062 result2062 = new Result2062();
								result2062.setSystem(StringUtils.StringExtract(book.getBookName()));
								result2062.setScreenId(row.getColumn2());
								result2062.setScreenName(row.getColumn3());
								result2062.setButtonId(null);
								result2062.setButtonName(null);
								result2062.setController(sheet.get(key).getSheetName().replace("カスタムタグ定義", ""));
								result2062.setPattern(null);
								result2062.setSheetName(sheet.get(key).getSheetName());
								result2062.setMenuItem(null);
								result2062.setOnclickContent(f.get(row).toString());
								resultList.add(result2062);
							}
						}
					}
				}
			}
		}
		WriteExcel.write(path, "\\report.xls", new String[] {"サブシステム","画面ID","画面名","ボタンID","ボタン名","コントロール","パターン","シート","メニュー項目Name","パターン_開発者向け","Onclick"}, 6, resultList);
	}
}
