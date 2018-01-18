package per.owisho.learn.util.excel.eventread;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

/**
 * 通用excel表单数据解析器
 * 
 * BSDataHandler 是使用该类的需要自己去实现接口，
 * 该接口用来要求调用方实现自己的业务数据处理
 * 
 * 参考文件地址： https://www.cnblogs.com/skyivben/archive/2007/09/23/903582.html
 * 比较明显的open xml内容 https://poi.apache.org/spreadsheet/how-to.html#xssf_sax_api
 * poi官网示例 http://blog.csdn.net/adairxy/article/details/51832282 open xml相关介绍
 * 
 * @author owisho
 * @version 1.0
 * @date 2018年1月18日
 */
@SuppressWarnings("rawtypes")
public class SheetDataPraser extends DefaultHandler {

	// --------------为进行数据处理存在的元素（begin）-------------------
	private String lastContents;

	private boolean nextIsString;
	// --------------为进行数据处理存在的元素（end）-------------------
	// TODO 待确认属性 官网例子中有
	private SharedStringsTable sst;

	// xml中数据
	private ArrayList[] datas;

	// 缓存能力-因为事件方式解析目的即为减少内存占用，所以定义容量信息，超过容量信息的内容需要进行处理
	private int capacity;

	// 当前标记位
	private int index;

	private ArrayList<Object> currentData;

	private BSDataHandler handler;

	public SheetDataPraser(SharedStringsTable sst) {
		this(sst, 10);
	}
	
	public SheetDataPraser(SharedStringsTable sst,int capacity) {
		this(sst, capacity, new BSDataHandler() {
			@Override
			public void process(List[] datas) {
				for (List list : datas) {
					if (list == null)
						continue;
					System.out.println(list);
				}
			}
		});
	}

	public SheetDataPraser(SharedStringsTable sst, int capacity, BSDataHandler handler) {
		this.sst = sst;
		this.capacity = capacity;
		this.datas = new ArrayList[capacity];
		this.index = 0;
		this.handler = handler;
	}

	/**
	 * index 初始值为0 数据增加时index+1当index = capacity时表示数据已满
	 * 
	 * @return
	 */
	public boolean isFull() {
		return index == capacity;
	}

	public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {
		if (name.equals("row")) {
			if (currentData == null) {
				currentData = new ArrayList<Object>();
				datas[index] = currentData;
				index++;
			} else {
				if (isFull()) {
					dealAndClearDatas(handler);
				}
				datas[index] = currentData;
				currentData = new ArrayList<Object>();
				index++;
			}
		}
		// c=>cell
		if (name.equals("c")) {
			// Print the cell reference
			// System.out.println(attributes.getValue("r")+"-");
			// Figure out if the value is an index in the SST
			String cellType = attributes.getValue("t");
			if (cellType != null && cellType.equals("s")) {
				nextIsString = true;
			} else {
				nextIsString = false;
			}
			// Clear contents cache
			lastContents = "";
		}
	}

	public void endElement(String uri, String localName, String name) throws SAXException {
		// Process the last contents as required.
		// Do new,as characters() may be called more than once
		if (nextIsString) {
			int idx = Integer.parseInt(lastContents);
			lastContents = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
			nextIsString = false;
		}

		// v=>contents of a cell
		// Output after we've seen the string contents
		if (name.equals("v")) {
			currentData.add(lastContents);
		}

		if (name.equals("sheetData")) {
			dealAndClearDatas(handler);
		}
	}

	public void characters(char[] ch, int start, int length) {
		lastContents += new String(ch, start, length);
	}

	public void dealAndClearDatas(BSDataHandler handler) {
		handler.process(datas);
		datas = new ArrayList[capacity];
		index = 0;
	}

}
