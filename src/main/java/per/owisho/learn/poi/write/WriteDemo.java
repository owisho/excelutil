package per.owisho.learn.poi.write;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;


/**
 * 学习SXSSF方式excel解析的学习类
 * @author wangyang
 *
 */
public class WriteDemo {
	
	private static final String fileNameForWindows = "C:\\Users\\owisho\\Desktop\\sxssf.xlsx";
	
	@SuppressWarnings("unused")
	private static final String fileNameForMac = "/Users/wangyang/Desktop/sxssf.xlsx";

	public static void main(String[] args) throws IOException {
		write();
		
//		XSSFWorkbook wb = new XSSFWorkbook();
//		XSSFSheet sheet = wb.createSheet();
//		Row row = sheet.createRow(0);
//		Cell cell = row.createCell(0);
//		cell.setCellValue("1");
//		XSSFSheet sheet2 = wb.createSheet();
//		Row row2 = sheet2.createRow(0);
//		Cell cell2 = row2.createCell(0);
//		cell2.setCellValue("1");
//		OutputStream os = new FileOutputStream(fileNameForWindows);
//		wb.write(os);
//		wb.close();
	}
	
	public static void write() throws IOException{
		SXSSFWorkbook wb = new SXSSFWorkbook(100);
		Sheet sh = wb.createSheet("sheet");
		for(int rownum=0;rownum<10000;rownum++) {
			Row row = sh.createRow(rownum);
			for(int cellnum = 0;cellnum<10;cellnum++) {
				Cell cell = row.createCell(cellnum);
				String address = new CellReference(cell).formatAsString();
				cell.setCellValue(address);
			}
		}
		//Rows with rownum < 900 are flushed and not accessible
//		for(int rownum= 0;rownum<900;rownum++) {
//			Assert.assertNull(sh.getRow(rownum));
//		}
		//ther last 100 rows are still in memory 
//		for(int rownum=900;rownum<1000;rownum++) {
//			Assert.assertNotNull(sh.getRow(rownum));
//		}
		
		FileOutputStream out = new FileOutputStream(fileNameForWindows);
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
		
		FileOutputStream out = new FileOutputStream(fileNameForWindows);
		wb.write(out);
		out.close();
		//dispose of temporary files backing this workbook on disk
		wb.dispose();
		wb.close();
	}
	
}
