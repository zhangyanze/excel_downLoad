package test;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.excel.ExcelDown;

public class Test {

	@org.junit.Test
	public void test(){
		List<Map<String,Object>> maps=new ArrayList<Map<String,Object>>();
		Map<String,Object> data=new HashMap<String, Object>();
		data.put("serialno", 123456);
		data.put("member_name", "����");
		data.put("name", "��Ҫ���");
		data.put("amount", 1000.00);
		maps.add(data);
		data=new HashMap<String, Object>();
		data.put("serialno", 45678);
		data.put("member_name", "����");
		data.put("name", "��ҪͶ��");
		data.put("amount", 5000.00);
		maps.add(data);
		ExcelDown.down(maps);
	}
}
