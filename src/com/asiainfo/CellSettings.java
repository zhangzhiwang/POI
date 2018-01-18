package com.asiainfo;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 单元格的一些设置</br>
 * 主要包括：</br>
 * 1、设置日期格式
 * 
 * @author zhangzhiwang
 * @date 2018年1月18日 下午1:49:38
 */
public class CellSettings {
	public static void main(String[] args) {
		OutputStream out = null;
		Workbook workbook = null;
		try {
			out = new FileOutputStream("/Users/zhangzhiwang/Documents/poi/CellSettings.xls");
			workbook = new HSSFWorkbook();
			Sheet sheet = workbook.createSheet();
			Row row = sheet.createRow(0);
			Cell cell0 = row.createCell(0);

			Date date = new Date();
			cell0.setCellValue(date);//直接放Date进去，出来的时数字，必须格式化
			

			workbook.write(out);
			System.out.println("OK");
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (out != null) {
				try {
					out.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
			if (workbook != null) {
				try {
					workbook.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}
}
