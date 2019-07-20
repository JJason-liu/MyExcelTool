package jasontest.mytools.utils;

public class MyUtil {
	/**
	 * 插入字符
	 * 
	 * @param src
	 * @param dec
	 * @param position
	 * @return
	 */
	public static String insertStringInParticularPosition(String src, String dec, int position) {
		StringBuffer stringBuffer = new StringBuffer(src);
		return stringBuffer.insert(position, dec).toString();
	}

	/**
	 * 判断是否全为数字
	 * 
	 * @param s
	 * @return
	 */
	public static boolean isNumeric(String s) {
		if (s != null && !"".equals(s.trim()))
			return s.matches("^[0-9]*$");
		else
			return false;
	}

	/**
	 * 判断时候全为汉字
	 * 
	 * @param s
	 * @return
	 */
	public static boolean isChina(String s) {
		String reg = "[\\u4e00-\\u9fa5]+";
		return s.matches(reg);
	}
}
