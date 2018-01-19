package com.asiainfo;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 编辑Excel（在原Excel中做修改）
 * 
 * @author zhangzhiwang
 * @date 2018年1月19日 下午3:37:09
 */
public class ModifyExcel {
	public static void main(String[] args) {
		// 需求：如果第一行第一列的值是“aa”，那么将第二行第二列的背景色改为绿色
		// 先读再改
		InputStream in = null;
		OutputStream out = null;
		Workbook wb = null;
		try {
			in = new FileInputStream("/Users/zhangzhiwang/Documents/poi/origin.xls");
			wb = new HSSFWorkbook(in);
			Sheet sheet = wb.getSheetAt(0);
			Row row0 = sheet.getRow(0);
			Cell cell0 = row0.getCell(0);
			if (cell0 != null && "aa".equals(cell0.getStringCellValue())) {
				Row row1 = sheet.getRow(1);
				if (row1 == null) {
					row1 = sheet.createRow(1);
				}
				Cell cell1 = row1.getCell(1);
				if (cell1 == null) {
					cell1 = row1.createCell(1);
					CellStyle cellStyle = wb.createCellStyle();
					cellStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
					cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
					cell1.setCellStyle(cellStyle);
				}
			}

			out = new FileOutputStream("/Users/zhangzhiwang/Documents/poi/origin.xls");
			wb.write(out);
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
			if (wb != null) {
				try {
					wb.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
			if (in != null) {
				try {
					in.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}
}
