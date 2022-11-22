package common;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class StringUtils {

	/**
	 * 提取「」中的文字
	 * @param string
	 * @return
	 */
	public static String StringExtract(String string) {
		String result = "";
		Pattern pattern = Pattern.compile("\\「(.*?)\\」");
		Matcher matcher = pattern.matcher(string);
		while(matcher.find()) {
			result = matcher.group(1);
		}
		return result;
	}
}
