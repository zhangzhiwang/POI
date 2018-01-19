package com.asiainfo;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * 单元格合并
 * 
 * @author zhangzhiwang
 * @date 2018年1月19日 下午2:37:00
 */
public class CellMerge {
	public static void main(String[] args) {
		OutputStream out = null;
		Workbook wb = null;
		try {
			out = new FileOutputStream("/Users/zhangzhiwang/Documents/poi/CellMerge.xls");
			wb = new HSSFWorkbook();
			Sheet sheet = wb.createSheet();
			Row row = sheet.createRow(0);
			Cell cell = row.createCell(0);
			cell.setCellValue("aaa");

			// 合并三行两列
			sheet.addMergedRegion(new CellRangeAddress(0, // firstRow起始行的索引（不是行数）
					2, // 结束行的索引
					0, // 起始列的索引
					1));// 结束列的索引

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
