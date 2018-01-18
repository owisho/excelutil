package per.owisho.learn.poi.write;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Assert;


/**
 * 学习SXSSF方式excel解析的学习类
 * @author wangyang
 *
 */
public class WriteDemo {

	public static void main(String[] args) throws IOException {
		write();
	}
	
	public static void write() throws IOException{
		SXSSFWorkbook wb = new SXSSFWorkbook(100);
		Sheet sh = wb.createSheet();
		for(int rownum=0;rownum<1000;rownum++) {
			Row row = sh.createRow(rownum);
			for(int cellnum = 0;cellnum<10;cellnum++) {
				Cell cell = row.createCell(cellnum);
				String address = new CellReference(cell).formatAsString();
				cell.setCellValue(address);
			}
		}
		//Rows with rownum < 900 are flushed and not accessible
		for(int rownum= 0;rownum<900;rownum++) {
			Assert.assertNull(sh.getRow(rownum));
		}
		//ther last 100 rows are still in memory 
		for(int rownum=900;rownum<1000;rownum++) {
			Assert.assertNotNull(sh.getRow(rownum));
		}
		
		FileOutputStream out = new FileOutputStream("/Users/wangyang/Desktop/sxssf.xlxs");
		wb.write(out);
		out.close();
		wb.dispose();
		wb.close();
	}
	
	public static void write1() throws IOException {
		SXSSFWorkbook wb = new SXSSFWorkbook(-1);//turn off auto-flushing and accumulate all rows in memory
		Sheet sh = wb.createSheet();
		for(int rownum = 0;rownum<1000;rownum++) {
			Row row = sh.createRow(rownum);
			for(int cellnum=0;cellnum<10;cellnum++) {
				Cell cell = row.createCell(cellnum);
				String address = new CellReference(cell).formatAsString();
				cell.setCellValue(address);
			}
			//manually control how rows are flushed to disk 
			if(rownum % 100 == 0) {
				((SXSSFSheet)sh).flushRows(100);//retain 100 last rows and flush all others
				//((SXSSFSheet)sh).flushRows() is shortcut for ((SXSSFSheet)sh).flushRows(0),
				//this method flushes all rows 
			}
		}
		
		FileOutputStream out = new FileOutputStream("/Users/wangyang/Desktop/sxssf.xlxs");
		wb.write(out);
		out.close();
		//dispose of temporary files backing this workbook on disk
		wb.dispose();
		wb.close();
	}
	
}
