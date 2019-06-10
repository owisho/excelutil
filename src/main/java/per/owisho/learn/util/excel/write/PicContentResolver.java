package per.owisho.learn.util.excel.write;

import lombok.extern.log4j.Log4j2;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import per.owisho.learn.util.IOUtil;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

@Log4j2
public class PicContentResolver implements ContentResolver{

	private Drawing patriarch;

	private int count = 0;

	private CellStyle cellStyle ;
	
	@Override
	public void resolve(Cell cell, Object content) {
		if(content==null){
			return;
		}
		if(!(content instanceof File)){
			throw new RuntimeException("内容不是文件");
		}
		File fileContent = (File)content;
		if(!fileContent.exists()){
			return;
		}
        try {
			count++;
			long start = System.currentTimeMillis();
			if(cellStyle==null){
				cellStyle = this.getContentStyle(cell.getSheet().getWorkbook());
			}
			dealPicture(cell, fileContent,2080, 20, 20 ,10 ,10 ,3 ,0.9 ,0.8,cellStyle);
			System.out.println("处理一张照片的时间："+(System.currentTimeMillis()-start)+"处理到的图片的张数："+count);
		} catch (IOException e) {
			System.out.println("图片文件地址："+fileContent.getAbsolutePath());
			log.error(e.getMessage(),e);
		}
		
	}

	/**
	 * @ Author wangfuliang
	 * 处理图片方法
	 * 一个单元格宽度分为1024份[0-1023]、高度分为256份[0-255]
	 * @param cell 当前单元格
	 * @param fileContent 图片流
	 * @param row_height 单元格高度
	 * @param columnWidth 单元格宽度
	 * @param dx1 图片左上角x坐标
	 * @param dy1 图片左上角y坐标
	 * @param dx2 图片右下角x坐标
	 * @param dy2 图片右下角y坐标
	 * @param resizeX 图片横向压缩比例
	 * @param resizeY 图片竖向压缩比例
	 * @param cellStyle 单元格属性
	 * @throws IOException
	 *
	 */
	private void dealPicture(Cell cell,File fileContent,int row_height,
							 int columnWidth ,int dx1,int dy1,int dx2,int dy2,
							 double resizeX ,double resizeY ,CellStyle cellStyle) throws IOException {
	    cell.getRow().setHeight((short)row_height);
		int column_width=256*columnWidth+184;
		cell.getSheet().setColumnWidth(cell.getColumnIndex(), column_width);
		cell.setCellStyle(cellStyle);
		ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
//		BufferedImage bufferImg = ImageIO.read(fileContent);
//		ImageIO.write(bufferImg, fileContent.getName().substring(fileContent.getName().lastIndexOf(".")+1), byteArrayOut);
        byte[] buffer = new byte[1024];
        IOUtil.bufferedOutput(buffer,byteArrayOut,new FileInputStream(fileContent));
		Workbook workbook = cell.getSheet().getWorkbook();
		if(patriarch == null) {
            patriarch = cell.getSheet().createDrawingPatriarch();
        }
		int columnIndex = cell.getColumnIndex();
		int rowIndex = cell.getRowIndex();
		XSSFClientAnchor anchor = new XSSFClientAnchor(XSSFShape.EMU_PER_PIXEL*dx1, XSSFShape.EMU_PER_PIXEL*dy1,
				XSSFShape.EMU_PER_PIXEL*dx2, XSSFShape.EMU_PER_PIXEL*dy2,
				(short)columnIndex, rowIndex, (short)columnIndex+1, rowIndex+1);
		anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_AND_RESIZE);
        int pic_index = workbook.addPicture(byteArrayOut.toByteArray(), XSSFWorkbook.PICTURE_TYPE_JPEG);
		Picture picture = patriarch.createPicture(anchor, pic_index);
        picture.resize(resizeX,resizeY);
	}

}
