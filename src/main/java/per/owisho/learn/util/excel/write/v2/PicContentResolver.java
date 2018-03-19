package per.owisho.learn.util.excel.write.v2;

import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;

import javax.imageio.ImageIO;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

public class PicContentResolver implements ContentResolver{

	private HSSFPatriarch patriarch;
	
	@Override
	public void resolve(HSSFCell cell, Object content) {
		assert content instanceof File : "内容不是文件";
		File fileContent = (File)content;
	
		cell.getRow().setHeight((short)(1000));
		cell.getSheet().setColumnWidth(cell.getColumnIndex(), (short)1000);
		
		BufferedImage bufferImg = null;
        try {
        	ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();     
        	bufferImg = ImageIO.read(fileContent);
			ImageIO.write(bufferImg, fileContent.getName().substring(fileContent.getName().lastIndexOf(".")+1), byteArrayOut);
			
			Workbook workbook = cell.getSheet().getWorkbook();
			if(patriarch == null) {
				patriarch = cell.getSheet().createDrawingPatriarch();
			}
			HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, 0, 0, (short)cell.getColumnIndex(), cell.getRowIndex(), (short)(cell.getColumnIndex()+1), cell.getRowIndex()+1);
			
			patriarch.createPicture(anchor, workbook.addPicture(byteArrayOut.toByteArray(), HSSFWorkbook.PICTURE_TYPE_JPEG));
		} catch (IOException e) {
			e.printStackTrace();
		}
		
	}
	
}
