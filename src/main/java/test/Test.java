package test;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Test {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		Pattern pattern = Pattern.compile("\\「(.*?)\\」");
		Matcher matcher = pattern.matcher("画面アイテム定義(カスタムタグ)Rev1.13_「A.製造指図管理」.xls");
		while(matcher.find()) {
			System.out.println(matcher.group(1));
		}
	}

}
