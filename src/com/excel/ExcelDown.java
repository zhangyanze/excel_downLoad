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
		String[] headers = {"�����", "�û���", "������", "�����(Ԫ)"};
		String[] filedNames = {"serialno", "member_name", "name", "amount"};
		ExportExcel<Map<String,Object>> exportExcel = new ExportExcel<Map<String,Object>>();
		//����excel�ļ����������뵽�ͻ���		  
		ByteArrayOutputStream stream = new ByteArrayOutputStream();
		exportExcel.exportExcel("���н��", headers,map, filedNames, stream);			
		InputStream excelStream = new ByteArrayInputStream(stream.toByteArray());	
		try {
			OutputStream output = new FileOutputStream("D:\\�ҵ����Ͽ�\\Documents\\Downloads\\excel.xls");
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
		//web����� ����
		/*String excelFileName = new String("���н���¼�б�.xls".getBytes("gb2312"), "ISO8859-1");
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
