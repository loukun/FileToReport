package entity;

import java.util.Map;

/**
 * Excel java对象
 * @author 30038737
 *
 */
public class ExcelBook {
	
	/**
	 * Excel文件名称
	 */
	private String bookName;
	
	private Map<String, ExcelSheet> excelSheets;

	public String getBookName() {
		return bookName;
	}

	public void setBookName(String bookName) {
		this.bookName = bookName;
	}

	public Map<String, ExcelSheet> getExcelSheets() {
		return excelSheets;
	}

	public void setExcelSheets(Map<String, ExcelSheet> excelSheets) {
		this.excelSheets = excelSheets;
	}

	@Override
	public String toString() {
		return "ExcelBook [bookName=" + bookName + ", excelSheets=" + excelSheets + "]";
	}

	
}
