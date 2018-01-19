package per.owisho.learn.util.excel.eventread.v1;

import java.io.InputStream;
import java.util.Iterator;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.XMLReaderFactory;

/**
 * excel 事件方式读取类，参考poi官网教程示例代码
 * 
 * @author owisho
 * @version 1.0
 * @date 2018年1月19日
 */
public class EventReader {

	// 内容解析器
	private int capacity;

	// 数据处理器
	private BSDataHandler bsDataHandler;

	public EventReader(int capacity, BSDataHandler bsDataHandler) {
		this.capacity = capacity;
		this.bsDataHandler = bsDataHandler;
	}

	/**
	 * 解析单个excel sheet页
	 * 
	 * @param fileName
	 *            文件名
	 * @param sheetIndex
	 *            表单页的序号，通常情况下表单页的序号从1开始
	 * @throws Exception
	 */
	public void processOneSheet(String fileName, Integer sheetIndex) throws Exception {
		OPCPackage pkg = OPCPackage.open(fileName);
		XSSFReader r = new XSSFReader(pkg);
		SharedStringsTable sst = r.getSharedStringsTable();

		XMLReader parser = fetchSheetParser(sst);

		InputStream sheet = r.getSheet("rId" + sheetIndex);
		InputSource sheetSource = new InputSource(sheet);
		parser.parse(sheetSource);
		sheet.close();
	}

	/**
	 * 解析整个excel文件（handler的处理方式没有完全想好，不推荐使用）
	 * 
	 * @param fileName
	 *            excel 文件名
	 * @throws Exception
	 */
	@Deprecated
	public void processAllSheets(String fileName) throws Exception {
		OPCPackage pkg = OPCPackage.open(fileName);
		XSSFReader r = new XSSFReader(pkg);
		SharedStringsTable sst = r.getSharedStringsTable();

		XMLReader parser = fetchSheetParser(sst);

		Iterator<InputStream> sheets = r.getSheetsData();
		while (sheets.hasNext()) {
			InputStream sheet = sheets.next();
			InputSource sheetSource = new InputSource(sheet);
			parser.parse(sheetSource);
			sheet.close();
		}

	}

	public XMLReader fetchSheetParser(SharedStringsTable sst) throws SAXException {
		XMLReader parser = XMLReaderFactory.createXMLReader("org.apache.xerces.parsers.SAXParser");
		ContentHandler handler = null;
		if (bsDataHandler == null) {
			handler = new SheetDataPraser(sst, capacity);
		} else {
			handler = new SheetDataPraser(sst, capacity, bsDataHandler);
		}
		parser.setContentHandler(handler);
		return parser;
	}

}
