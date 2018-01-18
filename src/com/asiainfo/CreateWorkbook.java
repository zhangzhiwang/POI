package com.asiainfo;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 创建工作簿
 * 
 * @author zhangzhiwang
 * @date 2017年11月14日
 */
public class CreateWorkbook {
	public static void main(String[] args) {
		Workbook workbook = new HSSFWorkbook();
		OutputStream out = null;
		try {
			out = new FileOutputStream("/Users/zhangzhiwang/Documents/poi/创建一个工作簿.xls");
			workbook.write(out);
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
		}
	}
}