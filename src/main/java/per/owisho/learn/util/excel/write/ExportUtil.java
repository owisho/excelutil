package per.owisho.learn.util.excel.write;

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

import per.owisho.learn.util.excel.write.vo.ExcelCellVO;
import per.owisho.learn.util.excel.write.vo.ExcelRowVO;
import per.owisho.learn.util.excel.write.vo.ExcelVO;
import per.owisho.learn.util.excel.write.vo.SheetContentVO;
import per.owisho.learn.util.excel.write.vo.SheetTitleVO;
import per.owisho.learn.util.excel.write.vo.SheetVO;

/**
 * 导出工具类
 * 
 * @author wangyang
 * @version 1.0
 * @date 2018年01月12日
 */
public class ExportUtil {

	/**
	 * 使用实体数据，定位数据，excel名称，表单名称，标题栏数据解析器，内容数据解析器构造要导出的数据的数据格式
	 * @param excelName excel名称
	 * @param sheetName 表单名称
	 * @param datas 数据集合
	 * @param cellLocation 数据定位及标题栏数据
	 * @param titleResolver 标题栏数据解析器
	 * @param contentResolver 内容数据解析器
	 * @return
	 * @throws Exception
	 */
	public ExcelVO constructData(String excelName,String sheetName,List<Object> datas,LinkedHashMap<String,String> cellLocation,
			ArrayList<ContentResolver> titleResolver,ArrayList<ContentResolver> contentResolver) throws Exception {
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
		
		ArrayList<ExcelCellVO> cells = new ArrayList<ExcelCellVO>();
		for(int i=0;i<titles.size();i++) {
			ContentResolver resolver = null;
			if(null!=titleResolver) 
				resolver = titleResolver.get(i);
			String title = titles.get(i);
			ExcelCellVO cell = new ExcelCellVO(title, null, resolver);
			cells.add(cell);
		}
		ExcelRowVO titleRow = new ExcelRowVO(cells);
		SheetTitleVO sheetTitleVO = new SheetTitleVO(Arrays.asList(titleRow));
		
		ArrayList<ExcelRowVO> rows = new ArrayList<ExcelRowVO>();
		if(null!=datas&&!datas.isEmpty()) {
			for(Object o:datas) {
				ExcelRowVO row = dataToRow(locations, contentResolver, o);
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
	 * @return 用数据转化出来的行信息
	 * @throws NoSuchFieldException
	 * @throws SecurityException
	 * @throws IllegalArgumentException
	 * @throws IllegalAccessException
	 */
	private ExcelRowVO dataToRow(ArrayList<String> locations,ArrayList<ContentResolver> contentResolvers,Object data) throws NoSuchFieldException, SecurityException, IllegalArgumentException, IllegalAccessException {
		ArrayList<ExcelCellVO> cells = new ArrayList<ExcelCellVO>();
		Class<?> cls = data.getClass();
		boolean hasResolver = true;
		if(contentResolvers==null)
			hasResolver = false;
		for (int i = 0; i < locations.size(); i++) {
			String location = locations.get(i);
			Field field = cls.getDeclaredField(location);
			field.setAccessible(true);
			Object value = field.get(data);
			ContentResolver resolver = null;
			if(hasResolver)
				resolver = contentResolvers.get(i);
			ExcelCellVO cell = new ExcelCellVO(value, null, resolver);
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
				if (title != null && title.getRows() != null && !title.getRows().isEmpty()) {
					List<ExcelRowVO> rows = title.getRows();
					rowIndex = addSheetRows(sh, rows, rowIndex);
				}

				SheetContentVO content = sheet.getContent();
				if (content != null && content.getRows() != null && !content.getRows().isEmpty()) {
					List<ExcelRowVO> rows = content.getRows();
					rowIndex = addSheetRows(sh, rows, rowIndex);
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
						c.setCellStyle(cell.getStyle());
						c.setCellValue(cell.getContent());
					}
				}
			}
		}
		return rowIndex;
	}

	public static void main(String[] args) throws FileNotFoundException {

//		test1();
//		test2();
		
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
	
	@SuppressWarnings("unused")
	private static void test2() {
		
		List<Object> arrayList = new ArrayList<Object>();
		DemoBean bean = new DemoBean();
		bean.setDate(new Date());
		bean.setEmail("owisho@126.com");
		bean.setId(1);
		bean.setName("owisho");
		for(int i=0;i<5;i++) {
			arrayList.add(bean);
		}
		LinkedHashMap<String, String> map = new LinkedHashMap<>();
		map.put("序号", "id");
		map.put("姓名","name");
		map.put("邮箱","email");
		map.put("日期","date");
		ExportUtil util = new ExportUtil();
		try {
			ExcelVO vo = util.constructData("test", "ceshi", arrayList, map, null, null);
			FileOutputStream os = new FileOutputStream("/usr/local/myImage/image/" + vo.getExcelName() + ".xls");
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
	
}
