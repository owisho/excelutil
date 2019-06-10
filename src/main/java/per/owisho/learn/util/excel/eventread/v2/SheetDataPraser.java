package per.owisho.learn.util.excel.eventread.v2;

import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import java.util.ArrayList;
import java.util.List;

/**
 * 通用excel表单数据解析器
 * 
 * BSDataHandler 是使用该类的需要自己去实现接口，
 * 该接口用来要求调用方实现自己的业务数据处理
 * 
 * 参考文件地址： https://www.cnblogs.com/skyivben/archive/2007/09/23/903582.html
 * poi官网示例 https://poi.apache.org/spreadsheet/how-to.html#xssf_sax_api
 * 比较明显的open xml内容 http://blog.csdn.net/adairxy/article/details/51832282 open xml相关介绍
 * 
 * @author owisho
 * @version 1.0
 * @date 2018年1月18日
 */
@SuppressWarnings("rawtypes")
public class SheetDataPraser extends DefaultHandler {

	// --------------为进行数据处理存在的元素（begin）-------------------
	private String lastContents;

	/**
	 * 当前单元格内容是否是字符串格式
	 */
	private boolean nextIsString;
	// --------------为进行数据处理存在的元素（end）-------------------
	// 所有共享字符串的数据表
	private SharedStringsTable sst;

	// xml中数据
	private ArrayList[] datas;

	// 缓存能力-因为事件方式解析目的即为减少内存占用，所以定义容量信息，超过容量信息的内容需要进行处理
	private int capacity;

	// 当前标记位
	private int index;

	private ArrayList currentData;

	private BSDataHandler handler;

	public SheetDataPraser(SharedStringsTable sst) {
		this(sst, 10);
	}

	public SheetDataPraser(SharedStringsTable sst,int capacity) {
		this(sst, capacity, new BSDataHandler() {
			@Override
			public void process(ArrayList[] datas) {
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

	private boolean isFirstRow = true;
	
	/**
	 * 读取到的当前行的最后一列
	 */
	private int lastcolumn = 0;
	
	/**
	 * 列标志；用来记录标题栏所具有的列字母（eg:ABCDE）
	 */
	private String colflags = "";
	
	@SuppressWarnings("unchecked")
	public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {
		if (name.equals("row")) {
			if (currentData == null) {
				currentData = new ArrayList();
				datas[index] = currentData;
				index++;
			} else {
				if (isFull()) {
					dealAndClearDatas(handler);
				}
				currentData = new ArrayList();
				datas[index] = currentData;
				index++;
			}
		}
		
		// c=>cell
		if (name.equals("c")) {
			
			//补充空元素逻辑--------start---------
			//将excel转化成xml格式的时候会将单元格的坐标转化成标签的r属性，xml会忽略为空的单元格
			String reference = attributes.getValue("r");
			//截取单元格的第一位，理论上26列以内的excel解析都可以正常进行，a-z
			String colflag = reference.substring(0,1);
			if(isFirstRow) {
				colflags += colflag;
			}else {
				//获取不为空的元素列
				int col = colflags.indexOf(colflag);
				//补上从当前行最后一列到不为空列之间缺少的列内容，用空数据填充
				for(int i=lastcolumn;i<col;i++) {
					if(currentData==null) {
						currentData = new ArrayList();
					}
					currentData.add(null);
					lastcolumn++;
				}
			}
			//补充空元素逻辑--------end-----------
			
			// Print the cell reference
			// System.out.println(attributes.getValue("r")+"-");
			String cellType = attributes.getValue("t");
			if (cellType != null && (cellType.equals("s")||cellType.equals("inlineStr"))) {
				nextIsString = true;
			} else {
				nextIsString = false;
			}
			// Clear contents cache 读取新内容的时候清空原来的内容
			lastContents = "";
		}
	}

	@SuppressWarnings("unchecked")
	public void endElement(String uri, String localName, String name) throws SAXException {
		//读取单元格内容----开始--------------
		// Process the last contents as required.
		// Do new,as characters() may be called more than once
		if (nextIsString) {
			try{
				int idx = Integer.parseInt(lastContents);
				lastContents = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
			}catch (Exception e){
			}
			nextIsString = false;
		}

		// v=>contents of a cell
		// Output after we've seen the string contents
//		if (name.equals("v")) {
		if(name.equals("c")){
			currentData.add(lastContents);
			//修改当前列的值
			lastcolumn++;
		}
		//读取单元格内容----结束---------------
		
		if (name.equals("row")) {
			//补充空元素逻辑--------start---------
			this.isFirstRow = false;
			if(lastcolumn<colflags.length()) {
				for(int i=lastcolumn;i<colflags.length();i++) {
					currentData.add(null);
				}
			}
			lastcolumn = 0;
			//补充空元素逻辑--------end---------
		}
		
		if (name.equals("sheetData")) {
			dealAndClearDatas(handler);
		}
	}

	/**
	 * 官方API示例中代码：具体原因未知，暂未找到源码
	 * 猜测功能，每个标签内容数据读取时会调用这个方法将数据写入lastContents中
	 */
	public void characters(char[] ch, int start, int length) {
		lastContents += new String(ch, start, length);
	}

	/**
	 * 将缓存中的数据进行相应的处理，然后清空缓存数据
	 * @param handler
	 */
	public void dealAndClearDatas(BSDataHandler handler) {
		handler.process(datas);
		datas = new ArrayList[capacity];
		index = 0;
	}

}
