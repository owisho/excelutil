package per.owisho.learn.util.excel.write.v2;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import per.owisho.learn.util.excel.write.v2.vo.ExcelCellVO;
import per.owisho.learn.util.excel.write.v2.vo.ExcelRowVO;
import per.owisho.learn.util.excel.write.v2.vo.ExcelVO;
import per.owisho.learn.util.excel.write.v2.vo.SheetContentVO;
import per.owisho.learn.util.excel.write.v2.vo.SheetTitleVO;
import per.owisho.learn.util.excel.write.v2.vo.SheetVO;

/**
 * 导出工具类
 * 
 * @author wangyang
 * @version 1.0
 * @date 2018年01月12日
 */
public class ExportUtil {

	/**
	 * TODO 待修改，行号应该是每个页签都是单独的
	 * 行号
	 */
	private int dataIndex = 1;
	
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
	public ExcelVO constructData(String excelName,String sheetName,List<Object> datas,LinkedHashMap<String,String> cellLocation,
			ArrayList<ContentResolver> titleResolver,ArrayList<ContentResolver> contentResolver,boolean addIndex) throws Exception {
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
		
		ContentResolver defaultResolver = DefaultResolverEnum.Title.getResolver();
		
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
		
		ContentResolver defaultResolver = DefaultResolverEnum.Content.getResolver();
		
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
	public void export(OutputStream os, ExcelVO vo) throws Exception {

		HSSFWorkbook wb = new HSSFWorkbook();
		try {
			List<SheetVO> sheets = vo.getSheets();
			if (sheets == null || sheets.isEmpty())
				throw new Exception("要导出的excel无内容");
			for (int i = 0; i < sheets.size(); i++) {
				SheetVO sheet = sheets.get(i);
				HSSFSheet sh = wb.createSheet(sheet.getSheetName() == null ? "sheet" + i : sheet.getSheetName());

				// 行号
				int rowIndex = 0;
				SheetTitleVO title = sheet.getTilte();
				//sheet页的列数
				int cols = -1;
				if (title != null && title.getRows() != null && !title.getRows().isEmpty()) {
					List<ExcelRowVO> rows = title.getRows();
					rowIndex = addSheetRows(sh, rows, rowIndex);
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
					rowIndex = addSheetRows(sh, rows, rowIndex);
					//如果没有标题行
					if(cols==-1) {
						cols = rows.get(0).getCells().size();
					}
				}
				
				//自动调整列宽待修改
				for(int j=0;j<cols;j++) {
					sh.autoSizeColumn(j);
				}
				
			}
			wb.write(os);
		} finally {
			wb.close();
		}
	}

	private Integer addSheetRows(HSSFSheet sh, List<ExcelRowVO> rows, Integer rowIndex) {
		if (rows != null && !rows.isEmpty()) {
			for (ExcelRowVO row : rows) {
				HSSFRow r = sh.createRow(rowIndex++);
				int cellIndex = 0;
				List<ExcelCellVO> cells = row.getCells();
				if (null != cells && !cells.isEmpty()) {
					for (ExcelCellVO cell : cells) {
						HSSFCell c = r.createCell(cellIndex++);
						cell.drawCell(c);
					}
				}
			}
		}
		return rowIndex;
	}

	public static void main(String[] args) throws FileNotFoundException {

//		test1();
		test2();
		
	}
	
	@SuppressWarnings("unused")
	private static void test1() throws FileNotFoundException {
		ExcelCellVO cell = new ExcelCellVO("1", null);
		List<ExcelCellVO> cells = new ArrayList<ExcelCellVO>();
		for (int i = 0; i < 5; i++) {
			cells.add(cell);
		}
		ExcelRowVO row = new ExcelRowVO(cells);
		List<ExcelRowVO> rows = new ArrayList<ExcelRowVO>();
		for (int i = 0; i < 10; i++) {
			rows.add(row);
		}
		SheetContentVO content = new SheetContentVO(rows);
		SheetVO sheet = new SheetVO("test", null, content);
		ExcelVO excel = new ExcelVO("ceshi", Arrays.asList(sheet));
		FileOutputStream os = new FileOutputStream("/usr/local/myImage/image/" + excel.getExcelName() + ".xls");
		ExportUtil util = new ExportUtil();
		try {
			util.export(os, excel);
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				os.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}
	
	private static void test2() {
		
		List<Object> arrayList = new ArrayList<Object>();
		DemoBean bean = new DemoBean();
		bean.setDate(new Date());
		bean.setEmail("owisho@126.com");
		bean.setName("owisho");
		bean.setFile(new File("C:\\Users\\owisho\\Desktop\\工作任务\\20180305\\file\\1.jpg"));
		for(int i=0;i<5;i++) {
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
			FileOutputStream os = new FileOutputStream("D:\\import\\tmp\\" + vo.getExcelName() + ".xls");
			try {
				util.export(os, vo);
			} catch (Exception e) {
				e.printStackTrace();
			} finally {
				try {
					os.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
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
