package com.excel;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;

public class ExcelDown {

	public static void down(List<Map<String,Object>> map){
		String[] headers = {"借款编号", "用户名", "借款标题", "借款金额(元)"};
		String[] filedNames = {"serialno", "member_name", "name", "amount"};
		ExportExcel<Map<String,Object>> exportExcel = new ExportExcel<Map<String,Object>>();
		//生成excel文件并将流输入到客户端		  
		ByteArrayOutputStream stream = new ByteArrayOutputStream();
		exportExcel.exportExcel("所有借款", headers,map, filedNames, stream);			
		InputStream excelStream = new ByteArrayInputStream(stream.toByteArray());	
		try {
			OutputStream output = new FileOutputStream("D:\\我的资料库\\Documents\\Downloads\\excel.xls");
			byte[] buf = new byte[1024];
			int len=0;
			while ((len = excelStream.read(buf)) > 0) {
				output.write(buf, 0, len);
			}
			excelStream.close();
			output.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
		//web浏览器 下载
		/*String excelFileName = new String("所有借款记录列表.xls".getBytes("gb2312"), "ISO8859-1");
		response.setContentType("application/x-msdownload");
		response.setCharacterEncoding("utf-8");
		response.setHeader("Content-disposition", "attachment; filename=" + excelFileName);
		OutputStream out = response.getOutputStream();
		byte[] buf = new byte[1024];
		int len=0;
		while ((len = excelStream.read(buf)) > 0) {
			out.write(buf, 0, len);
		}
		excelStream.close();
		out.close();*/
	}
}
