package per.owisho.learn.util.excel.eventread.v1.demo;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import per.owisho.learn.util.excel.eventread.v1.BSDataHandler;

/**
 * BSDataHandler的示例实现类
 * @author owisho
 * @version 1.0
 * @date 2018年01月19日
 */
public class DemoBSDataHandler implements BSDataHandler{

	private ArrayList<String> titles;
	
	private ArrayList<String> attributes;
	
	public DemoBSDataHandler(ArrayList<String> titles,ArrayList<String> attributes) {
		this.titles = titles;
		this.attributes = attributes;
	}
	
	@SuppressWarnings({ "rawtypes", "unchecked" })
	@Override
	public void process(List[] datas) {
		if(null!=datas&&datas.length>0) {
			for(List data:datas) {
				if(data==null)
					continue;
				if(data.containsAll(titles)&&titles.containsAll(data))
					continue;
				if(data.size()==attributes.size()) {
					Map<String,Object> map = new HashMap<String,Object>();
					for(int i=0;i<data.size();i++) {
						map.put(attributes.get(i), data.get(i));
					}
					System.out.println(map);
				}
			}
		}
	}

}
