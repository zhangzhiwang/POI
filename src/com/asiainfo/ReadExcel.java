package com.asiainfo;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;

/**
 * 读取Excel
 *
 * @author zhangzhiwang
 * @date 2018年1月18日 下午8:48:07
 */
public class ReadExcel {
	public static void main(String[] args) {
		InputStream in = null;
		try {
			in = new FileInputStream("/Users/zhangzhiwang/Documents/poi/ReadExcel.xls");
			HSSFWorkbook wk = new HSSFWorkbook(in);
			HSSFSheet sheet = wk.getSheetAt(0);// 获取第一个工作表
			// 应该判空，此处为了演示，故没有严谨
			for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++) {//rowNum即行数是从0开始的，可以理解为行数就是行的索引，从0开始
				HSSFRow row = sheet.getRow(rowIndex);
				for (int cellIndex = 0; cellIndex < row.getLastCellNum(); cellIndex++) {//cellNum即列数是从1开始的，和rowNum不一样（很恶心是不是），所以cellNum不能理解为列索引
					HSSFCell cell = row.getCell(cellIndex);//总结：rowIndex和cellIndex都是从0开始；rowNum从0开始，cellNum从1开始
					System.out.print(getCellValue(cell) + "|");
				}
				System.out.println();
			}

		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (in != null) {
				try {
					in.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}

	private static Object getCellValue(HSSFCell cell) {
		if (cell.getCellTypeEnum() == CellType.NUMERIC) {//如何读取时间格式？
			return cell.getNumericCellValue();
		} else if (cell.getCellTypeEnum() == CellType.STRING) {
			return cell.getStringCellValue();
		} else if (cell.getCellTypeEnum() == CellType.BOOLEAN) {
			return cell.getBooleanCellValue();
		}
		return null;
	}
}
