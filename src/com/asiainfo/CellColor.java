package com.asiainfo;

import java.io.FileOutputStream;
import java.io.IOException;
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
 * 单元格颜色
 * 
 * @author zhangzhiwang
 * @date 2018年1月19日 下午2:26:24
 */
public class CellColor {
	public static void main(String[] args) {
		OutputStream out = null;
		Workbook wb = null;
		try {
			out = new FileOutputStream("/Users/zhangzhiwang/Documents/poi/CellColor.xls");
			wb = new HSSFWorkbook();
			Sheet sheet = wb.createSheet();
			Row row = sheet.createRow(0);
			Cell cell = row.createCell(0);
			cell.setCellValue("aaa");

			CellStyle cellStyle = wb.createCellStyle();
			//前景色（p.s.前景色和背景色有什么区别？）
			cellStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
			cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			
			Cell cell2 = row.createCell(1);
			cell2.setCellValue("123");
			CellStyle cellStyle2 = wb.createCellStyle();
			//背景色（目前测试不起作用）
			cellStyle2.setFillBackgroundColor(IndexedColors.RED.getIndex());
			cellStyle2.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			cell.setCellStyle(cellStyle);
			wb.write(out);
			System.out.println("OK");
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (wb != null) {
				try {
					wb.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
			if (out != null) {
				try {
					out.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}
}
