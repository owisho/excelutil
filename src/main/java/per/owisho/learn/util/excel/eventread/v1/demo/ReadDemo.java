package per.owisho.learn.util.excel.eventread.v1.demo;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;

import org.apache.poi.hssf.eventusermodel.HSSFEventFactory;
import org.apache.poi.hssf.eventusermodel.HSSFRequest;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import per.owisho.learn.poi.read.EventExample;
import per.owisho.learn.util.excel.eventread.v1.EventReader;

public class ReadDemo {

	public static void main(String[] args) throws Exception {
//		System.out.println(args[0]);
		readForXLXS(args[0], 1);
	}

	/**
	 * 解析xls型文档 Read an excel file and spit out what we find
	 * 
	 * @param args
	 *            Expect one argument that is the file to read.
	 * @throws IOException
	 *             When there is an error processing the file
	 */
	public static void readForXLS(String fileName) throws IOException {

		// create a new file inout stream with the input file specified
		// at the command line
		FileInputStream fin = new FileInputStream(fileName);
		// create a new org.apache.poi.poifs.filesystem.Filesystem
		@SuppressWarnings("resource")
		POIFSFileSystem poifs = new POIFSFileSystem(fin);
		// get the Workbook(excel part) stream in a InputStream
		InputStream din = poifs.createDocumentInputStream("Workbook");
		// construct out HSSFRequest object
		HSSFRequest req = new HSSFRequest();
		// lazy listen for all records with the listener shown above
		req.addListenerForAllRecords(new EventExample());
		// create our event factory
		HSSFEventFactory factory = new HSSFEventFactory();
		// process our events based on the document input stream
		factory.processEvents(req, din);
		// once all the events are processed close our file input stream
		fin.close();
		// and our document input stream (don't want to leak these!)
		din.close();
		System.out.println("done.");
	}

	public static void readForXLXS(String fileName,Integer sheetIndex) throws Exception {
		ArrayList<String> titles = new ArrayList<String>(5);
		titles.add("标题1");
		titles.add("标题2");
		titles.add("标题3");
		titles.add("标题4");
		titles.add("标题5");
		
		ArrayList<String> attributes = new ArrayList<String>(5);
		attributes.add("attributes1");
		attributes.add("attributes2");
		attributes.add("attributes3");
		attributes.add("attributes4");
		attributes.add("attributes5");
		DemoBSDataHandler demoHandler = new DemoBSDataHandler(titles,attributes);
		EventReader reader = new EventReader(10,demoHandler);
		reader.processOneSheet(fileName, sheetIndex);
	}
	
}
