package per.owisho.learn.util.excel.write;

import lombok.extern.log4j.Log4j2;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import per.owisho.learn.util.excel.write.vo.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.*;
import java.util.concurrent.CountDownLatch;

/**
 * 导出工具类
 *
 * @author wangyang
 * @version 1.0
 * @date 2018年01月12日
 */
@Log4j2
public class ExportUtil {

	/**
	 * TODO 待修改，行号应该是每个页签都是单独的
	 * 行号
	 */
	private int dataIndex = 1;

	private int flushCacheSize = 30;

	private int dealnum = 0;

	/**
	 * 使用实体数据，定位数据，excel名称，表单名称，标题栏数据解析器，内容数据解析器构造要导出的数据的数据格式
	 * @param excelName excel名称
	 * @param sheetName 表单名称
	 * @param datas 数据集合
	 * @param addIndex 是否添加序号
	 * @param cellLocation 数据定位及标题栏数据
	 * @param titleResolver 标题栏数据解析器
	 * @param contentResolver 内容数据解析器
	 * @return
	 * @throws Exception
	 */
	public ExcelVO constructData(String excelName, String sheetName, List<?> datas, LinkedHashMap<String,String> cellLocation,
								 ArrayList<ContentResolver> titleResolver, ArrayList<ContentResolver> contentResolver, boolean addIndex) throws Exception {
		long start = System.currentTimeMillis();
		assert cellLocation!=null:"定位信息不能为空";
		if(titleResolver!=null)
			assert cellLocation.size()==titleResolver.size():"标题栏处理器数量不正确";
		if(contentResolver!=null)
			assert cellLocation.size()==contentResolver.size():"内容处理数量不正确";

		ArrayList<String> titles = new ArrayList<String>(cellLocation.size());
		ArrayList<String> locations = new ArrayList<String>(cellLocation.size());
		cellLocation.forEach((key,value)->{
			titles.add(key);
			locations.add(value);
		});

		ContentResolver defaultResolver = DefaultResolver.getTitleResolver();

		ArrayList<ExcelCellVO> cells = new ArrayList<ExcelCellVO>();

		if(addIndex) {
			ExcelCellVO cell = new ExcelCellVO("序号",defaultResolver);
			cells.add(cell);
		}
		for(int i=0;i<titles.size();i++) {
			ContentResolver resolver = (null==titleResolver)||(titleResolver.get(i)==null)?defaultResolver:titleResolver.get(i);
			String title = titles.get(i);
			ExcelCellVO cell = new ExcelCellVO(title, resolver);
			cells.add(cell);
		}
		ExcelRowVO titleRow = new ExcelRowVO(cells);
		SheetTitleVO sheetTitleVO = new SheetTitleVO(Arrays.asList(titleRow));

		ArrayList<ExcelRowVO> rows = new ArrayList<ExcelRowVO>();
		if(null!=datas&&!datas.isEmpty()) {
			for(Object o:datas) {
				ExcelRowVO row = dataToRow(locations, contentResolver, o,addIndex);
				rows.add(row);
			}
		}
		SheetContentVO sheetContentVO = new SheetContentVO(rows);

		SheetVO sheetVO = new SheetVO(sheetName, sheetTitleVO, sheetContentVO);

		ExcelVO excel = new ExcelVO(excelName, Arrays.asList(sheetVO));

		System.out.println("ExportUtil.constructData"+(System.currentTimeMillis()-start));
		return excel;

	}

	/**
	 * 将实体数据转化成Excel行数据，本方法没有各种校验
	 * 认为调用方法前内容已经做了相关验证
	 * @param locations 定位列表
	 * @param contentResolvers 内容处理器列表
	 * @param data 实体数据
	 * @param addIndex 是否添加序号
	 * @return 用数据转化出来的行信息
	 * @throws NoSuchFieldException
	 * @throws SecurityException
	 * @throws IllegalArgumentException
	 * @throws IllegalAccessException
	 */
	private ExcelRowVO dataToRow(ArrayList<String> locations,ArrayList<ContentResolver> contentResolvers,Object data,boolean addIndex) throws NoSuchFieldException, SecurityException, IllegalArgumentException, IllegalAccessException {

		ContentResolver defaultResolver = DefaultResolver.getContentResolver();

		ArrayList<ExcelCellVO> cells = new ArrayList<ExcelCellVO>();
		Class<?> cls = data.getClass();
		if(addIndex) {
			ExcelCellVO cell = new ExcelCellVO(dataIndex++,defaultResolver);
			cells.add(cell);
		}
		for (int i = 0; i < locations.size(); i++) {

			String location = locations.get(i);
			Field field = cls.getDeclaredField(location);
			field.setAccessible(true);
			Object value = field.get(data);
			ContentResolver resolver = (contentResolvers==null)||(contentResolvers.get(i)==null)?defaultResolver:contentResolvers.get(i);

			ExcelCellVO cell = new ExcelCellVO(value, resolver);
			cells.add(cell);
		}
		ExcelRowVO row = new ExcelRowVO(cells);
		return row;
	}

	/**
	 * 导出数据方法，os需要调用方法自行处理（关闭和创建）
	 * @param os
	 * @param vo
	 * @throws Exception
	 */
	public void export(OutputStream os, ExcelVO vo,Long taskId) throws Exception {
		long start = System.currentTimeMillis();
		SXSSFWorkbook wb = new SXSSFWorkbook(null,200,true,false);
		try {
			List<SheetVO> sheets = vo.getSheets();
			if (sheets == null || sheets.isEmpty())
				throw new Exception("要导出的excel无内容");
			for (int i = 0; i < sheets.size(); i++) {
				SheetVO sheet = sheets.get(i);
				SXSSFSheet sh = wb.createSheet(sheet.getSheetName() == null ? "sheet" + i : sheet.getSheetName());

				// 行号
				int rowIndex = 0;
				SheetTitleVO title = sheet.getTilte();
				//sheet页的列数
				int cols = -1;
				if (title != null && title.getRows() != null && !title.getRows().isEmpty()) {
					List<ExcelRowVO> rows = title.getRows();
					rowIndex = addSheetRows(sh, rows, rowIndex,taskId);
					//使用最大标题的列的大小定义整个sheet页的列数
					for(ExcelRowVO rowVO :rows) {
						if(rowVO.getCells().size()>cols) {
							cols = rowVO.getCells().size();
						}
					}
				}

				SheetContentVO content = sheet.getContent();
				if (content != null && content.getRows() != null && !content.getRows().isEmpty()) {
					List<ExcelRowVO> rows = content.getRows();

					if(rows==null||rows.isEmpty())
						continue;

					//数据量小的时候单线程执行
//					if(rows.size()<1000){
					rowIndex = addSheetRows(sh, rows, rowIndex,taskId);
//					}
					//如果数据条数多余10000条
					/*
					else{
						int threadNum = 8;
						ThreadPoolCommon pool = new ThreadPoolCommon(threadNum);
						CountDownLatch doneSignal = new CountDownLatch(threadNum);
						int minSize = rows.size()/threadNum;
						for(int j=0;j<(threadNum-1);j++){
							List<ExcelRowVO> subRows = new ArrayList<>();
							for(int k=j*minSize;k<(j+1)*minSize;k++){
								subRows.add(rows.get(k));
							}
							pool.execute(new addSheetRowsTask(sh,subRows,j*minSize+rowIndex,doneSignal));
						}
						List<ExcelRowVO> subRows = new ArrayList<>();
						for(int j=(threadNum-1)*minSize;j<rows.size();j++){
							subRows.add(rows.get(j));
						}
						pool.execute(new addSheetRowsTask(sh,subRows,(threadNum-1)*minSize+rowIndex,doneSignal));
						doneSignal.await();
						pool.shutdown();
					}
					*/
					//如果没有标题行
					if(cols==-1) {
						cols = rows.get(0).getCells().size();
					}

				}

//				for(int num=0;num<cols;num++){
//					sh.trackAllColumnsForAutoSizing();
//					sh.autoSizeColumn(num);
//				}

			}
			//TODO 导出图片时excel可能会显示有点问题
			ZipSecureFile.setMinInflateRatio(0.005);
			wb.write(os);
		} finally {
			wb.close();
			System.out.println("ExportUtil.export"+(System.currentTimeMillis()-start));
		}
	}

	private class addSheetRowsTask implements Runnable{
		private SXSSFSheet sh ;
		private List<ExcelRowVO> rows;
		private int rowIndex;
		private CountDownLatch doneSignal;
		private long taskId;
		addSheetRowsTask(SXSSFSheet sh, List<ExcelRowVO> rows, int rowIndex,CountDownLatch doneSignal,Long taskId){
			this.sh = sh;
			this.rows = rows;
			this.rowIndex = rowIndex;
			this.doneSignal = doneSignal;
			this.taskId = taskId;
		}

		@Override
		public void run() {
			try{
				addSheetRows(sh,rows,rowIndex,taskId);
			}catch (Exception e){
                log.error(e.getMessage(),e);
			}finally {
				doneSignal.countDown();
			}
		}
	}

	private Integer addSheetRows(SXSSFSheet sh, List<ExcelRowVO> rows, int rowIndex,Long taskId) {

		//如果无数据直接返回
		if(rows==null||rows.isEmpty())
			return rowIndex;

		if (rows != null && !rows.isEmpty()) {
			for (ExcelRowVO row : rows) {
				SXSSFRow r = sh.createRow(rowIndex++);
				int cellIndex = 0;
				List<ExcelCellVO> cells = row.getCells();
				if (null != cells && !cells.isEmpty()) {
					for (ExcelCellVO cell : cells) {
						SXSSFCell c = r.createCell(cellIndex++);
						cell.drawCell(c);
					}
				}
				if(sh.getPhysicalNumberOfRows()%flushCacheSize==0){
					try {
						System.out.println("将数据全部刷新到硬盘");
						sh.flushRows();
						System.out.println("冲刷数据结束");
					} catch (IOException e) {
                        log.error(e.getMessage(),e);
					}
				}
			}
		}
		//TODO 数据全部处理完，冲刷数据
        try {
            sh.flushRows();
        } catch (IOException e) {
            log.error(e.getMessage(),e);
        }
        return rowIndex;
	}

	public static void main(String[] args) {
	    int rownum = 1000;
	    try{
	        rownum = Integer.parseInt(args[0]);
        }catch (Exception e){
            log.error(e.getMessage(),e);
        }
		test2(rownum);
	}

	private static void test2(int rownum) {

		List<Object> arrayList = new ArrayList<Object>();
		DemoBean bean = new DemoBean();
		bean.setDate(new Date());
		bean.setEmail("owisho@126.com");
		bean.setName("owisho");
		bean.setFile(new File("/mnt/data/wy/1.jpg"));
		for(int i=0;i<rownum;i++) {
			arrayList.add(bean);
		}
		LinkedHashMap<String, String> map = new LinkedHashMap<>();
		map.put("姓名","name");
		map.put("邮箱","email");
		map.put("日期","date");
		map.put("图片","file");

		ArrayList<ContentResolver> contentResolvers = new ArrayList<>();
		contentResolvers.add(null);
		contentResolvers.add(null);
		contentResolvers.add(null);
		contentResolvers.add(new PicContentResolver());

		ExportUtil util = new ExportUtil();
		try {
			ExcelVO vo = util.constructData("test", "ceshi", arrayList, map, null, contentResolvers,true);
			FileOutputStream os = new FileOutputStream("/mnt/data/wy" + vo.getExcelName() + ".xlsx");
			try {
				util.export(os, vo,null);
			} catch (Exception e) {
                log.error(e.getMessage(),e);
			} finally {
				try {
					os.close();
				} catch (IOException e) {
                    log.error(e.getMessage(),e);
				}
			}
		} catch (Exception e) {
            log.error(e.getMessage(),e);
		}

	}


}

class DemoBean {

	private Integer id ;

	private String name;

	private String email;

	private Date date;

	private File file;

	public Integer getId() {
		return id;
	}

	public void setId(Integer id) {
		this.id = id;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getEmail() {
		return email;
	}

	public void setEmail(String email) {
		this.email = email;
	}

	public Date getDate() {
		return date;
	}

	public void setDate(Date date) {
		this.date = date;
	}

	public File getFile() {
		return file;
	}

	public void setFile(File file) {
		this.file = file;
	}

}

