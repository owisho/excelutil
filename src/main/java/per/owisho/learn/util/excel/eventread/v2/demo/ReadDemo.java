package per.owisho.learn.util.excel.eventread.v2.demo;

import java.util.ArrayList;

import per.owisho.learn.util.excel.eventread.v2.EventReader;

public class ReadDemo {

	private static final String fileName = "C:\\Users\\owisho\\Desktop\\相机上传模板.xlsx";
	
	public static void main(String[] args) throws Exception {
//		System.out.println(args[0]);
		readForXLXS(fileName, 1);
	}

	public static void readForXLXS(String fileName,Integer sheetIndex) throws Exception {
		ArrayList<String> titles = new ArrayList<String>(5);
		titles.add("抓拍地址");
		titles.add("部署方案id");
		titles.add("lat");
		titles.add("lng");
		titles.add("设备位置");
		titles.add("设备名称");
		titles.add("备注");
		titles.add("视频流地址");
		titles.add("设备类型");
		titles.add("所属组织id");
		
		ArrayList<String> attributes = new ArrayList<String>(5);
		attributes.add("capture_src");
		attributes.add("deploy_solution_id");
		attributes.add("lat");
		attributes.add("lng");
		attributes.add("location");
		attributes.add("name");
		attributes.add("remark");
		attributes.add("src");
		attributes.add("type");
		attributes.add("user_group_id");
		
		DemoBSDataHandler demoHandler = new DemoBSDataHandler(titles,attributes);
		EventReader reader = new EventReader(10,demoHandler);
		reader.processOneSheet(fileName, sheetIndex);
	}
	
}
