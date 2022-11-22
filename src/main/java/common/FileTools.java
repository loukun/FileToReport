package common;

import java.io.File;
import java.util.List;

public class FileTools {
	
	/**
	 * 获取文件夹下的所有文件
	 * @param path
	 * @param outFileList
	 */
	public static void getAllFile(String path, List<File> outFileList) {
		
		File dirPath = new File(path);
		File[] fileList = dirPath.listFiles();
		assert fileList != null;
		for (File file : fileList) {
			if (file.isDirectory()) {
				getAllFile(file.getPath(), outFileList);
			} else {
				outFileList.add(file);
			}
		}
	}
}
