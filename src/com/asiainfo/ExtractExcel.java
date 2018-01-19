package com.asiainfo;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.extractor.ExcelExtractor;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 * 提取excel内容
 * 
 * @author zhangzhiwang
 * @date 2018年1月19日 下午1:04:45
 */
public class ExtractExcel {
	public static void main(String[] args) {
		InputStream in = null;
		HSSFWorkbook wb = null;
		try {
			in = new FileInputStream("/Users/zhangzhiwang/Documents/poi/ReadExcel.xls");
			wb = new HSSFWorkbook(in);
			ExcelExtractor excelExtractor = new ExcelExtractor(wb);
			excelExtractor.setIncludeSheetNames(true);
			System.out.println(excelExtractor.getText());
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if(wb != null) {
				try {
					wb.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
			if(in != null) {
				try {
					in.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		
	}
}
