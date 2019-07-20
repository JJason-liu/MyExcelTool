package jasontest.mytools;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import jasontest.mytools.utils.MyUtil;

public class Mytool {
	private static final String filePath = "D:\\Documents\\MuMu共享文件夹";
	private static final String path = filePath + "\\蓉城叶式族谱转换日期后文件"; 
	
	private static int count = 0;
	// 改变列表
//	private static final int[] changeArray = new int[] { 8, 10, 16, 18 };
	private static final int[] changeArray = new int[] { 7,8, 9,10, 15,16 ,17,18 };

	// 中文字符串
	private static final String[] chinaArray = new String[] { "〇", "一", "二", "三", "四", "五", "六", "七", "八", "九" };

	public static void main(String[] args) throws IOException {
		Mytool.readFiles(filePath);
		System.out.println("共计处理文件：" + count +"个");
	}

	public Mytool() {
	}

	/**
	 * 获取指定目录或者文件
	 * 
	 * @param path
	 */
	public static void readFiles(String path) throws IOException {
		File file = new File(path);
		if (file.isFile() && (file.getName().endsWith(".xls") || file.getName().endsWith(".xlsx"))) {
			System.out.println("开始处理文件：" + file.getName());
//			System.out.println(file.getAbsolutePath() + "********");
			file = copyFile(file);//不加这个，源文件会被修改，暂时添加，待思考新的方法
			readAndWriteFile(file);
			count++;
		} else if (file.isDirectory() && !file.getPath().contains("蓉城叶式族谱转换日期后文件")) {
			// System.out.println("目录名称： " + file.getName());
			String[] list = file.list();
			for (String string : list) {
				string = path + "\\" + string;
				// System.out.println(string);
				readFiles(string);
			}
		}
	}

	private static File copyFile(File file) {
		String newPath = path + "\\" +file.getName(); 
		File directfile = new File(path);
		// 判断是否存在
		if (!directfile.exists()) {
			// 创建父目录文件
			directfile.mkdirs();
		}

		FileInputStream fis = null;
		FileOutputStream fos = null;
		try {
			fis = new FileInputStream(file.getPath());
			fos = new FileOutputStream(newPath);
			// 创建搬运工具
			byte datas[] = new byte[1024 * 8];
			// 创建长度
			int len = 0;
			// 循环读取数据
			while ((len = fis.read(datas)) != -1) {
				fos.write(datas, 0, len);
			}
			fos.flush();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			try {
				// 3.释放资源
				fis.close();
				fos.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		return new File(newPath);
	}

	/**
	 * 读取单个文件
	 * 
	 * @param path
	 * @throws IOException
	 */
	public static void readAndWriteFile(File file) throws IOException {
		Workbook wb = null;
		String absolutePath = file.getAbsolutePath();
		if (file.getName().endsWith(".xls")) {// 2003版本
			FileInputStream fis = new FileInputStream(absolutePath); // 文件流对象
			wb = new HSSFWorkbook(fis);
		} else if (file.getName().endsWith(".xlsx")) {// 2007以上版本
			wb = new XSSFWorkbook(absolutePath);
		}

		Sheet sheet = wb.getSheetAt(0); // 读取sheet 0
//		System.out.println("当前处理的sheet页：" + sheet.getSheetName());

		int firstRowIndex = sheet.getFirstRowNum() + 3; // (第三列开始)
		int lastRowIndex = sheet.getLastRowNum();
//		System.out.println("firstRowIndex: " + firstRowIndex);
//		System.out.println("lastRowIndex: " + lastRowIndex);

		for (int rIndex = firstRowIndex; rIndex <= lastRowIndex; rIndex++) { // 遍历行
			Row row = sheet.getRow(rIndex);
			if (row != null) {
				// int firstCellIndex = row.getFirstCellNum();
				// int lastCellIndex = row.getLastCellNum();
				for (int cIndex : changeArray) { // 遍历列
					Cell cell = row.getCell(cIndex);
					String info = cell.toString();
					if (cell != null && !info.isEmpty() && !"?".equals(info)) {
//						System.out.println("rIndex: " + rIndex + "*************info: " + info);
						String changeChina = changeChina(info);
						cell.setCellValue(changeChina);
//						System.out.println("*************changeAfterinfo:  " + changeChina);
					}
				}
			}
		}
		writeFile(wb, file);
	}
 
	// 写文件
	private static void writeFile(Workbook wb, File file) {
		try {
			wb.write(new FileOutputStream(file.getName()));
			wb.close();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	private static String changeChina(String info) {
		String temp = "";
		for (int i = 0; i < info.length(); i++) {
			char charAt = info.charAt(i);
			String str = String.valueOf(charAt);
			if (Character.isDigit(charAt)) {// 只转换数字
				str = chinaArray[Integer.parseInt(str)];
			}
			temp += str;
		}
		return finalChange(temp);
	}

	public static String finalChange(String str) {
		if (MyUtil.isChina(str) || str.contains("不详") || str.length() < 3 || str.contains("日")) {
			return str;
		}

		if (str.length() == 4) {
			return str + "年";
		}
		String temp = "";
		String[] split = str.split("-");
		if (split.length <= 1) {
			temp = str;
		}
		for (int i = split.length - 1; i >= 0; i--) {
			temp += split[i];

		}
		// 添加年 月
		if (!"年".equals(temp.charAt(4) + "") && !"年".equals(temp.charAt(3) + "")) {
			temp = MyUtil.insertStringInParticularPosition(temp, "年", 4);
		}

		// 整理月字后面内容
		String[] split2 = temp.split("月");
		if (split2.length == 2 && split2[1].length() == 2) {
			temp = temp.substring(0, temp.indexOf("月") + 1);// 去掉月字后面的字
			String string = split2[1];
			if (string.startsWith("〇")) {// 排除一九五七年七月〇四
				string = string.substring(1, 2);
			} else {
				string = MyUtil.insertStringInParticularPosition(string, "十", 1);
				if (string.endsWith("〇")) {
					string = string.substring(0, string.indexOf("〇"));// 排除二〇〇四年一二十〇
				}
			}
			temp += string;
		}
		return temp + "日";
	}
}
