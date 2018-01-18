package per.owisho.learn.poi.read;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.eventusermodel.HSSFEventFactory;
import org.apache.poi.hssf.eventusermodel.HSSFRequest;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class ReadDemo {

	/**
	 * Read an excel file and spit out what we find 
	 * @param args Expect one argument that is the file to read.
	 * @throws IOException When there is an error processing the file
	 */
	public static void main(String[] args) throws IOException{
		//create a new file inout stream with the input file specified 
		//at the command line 
		FileInputStream fin = new FileInputStream(args[0]);
		//create a new org.apache.poi.poifs.filesystem.Filesystem
		@SuppressWarnings("resource")
		POIFSFileSystem poifs = new POIFSFileSystem(fin);
		//get the Workbook(excel part) stream in a InputStream
		InputStream din = poifs.createDocumentInputStream("Workbook");
		//construct out HSSFRequest object 
		HSSFRequest req = new HSSFRequest();
		//lazy listen for all records with the listener shown above
		req.addListenerForAllRecords(new EventExample());
		//create our event factory
		HSSFEventFactory factory = new HSSFEventFactory();
		//process our events based on the document input stream
		factory.processEvents(req, din);
		//once all the events are processed close our file input stream 
		fin.close();
		//and our document input stream (don't want to leak these!)
		din.close();
		System.out.println("done.");
	}
	
}
