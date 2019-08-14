package xmlchai;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.io.FileInputStream;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @FileName SplitExcelUtil.java
 * @Description: 分割文件excel工具类，将excel分割成多个
 */
public class SplitExcelUtil {
	/**
	 * @Title: splitExcels
	 * @Description:切割Excel文件方法
	 * @param filePath
	 * @return
	 */

	private static final String EXCEL_XLS = "xls";
	private static final String EXCEL_XLSX = "xlsx";
	// 每批导入excel总行数
	private static final int FileMaxNum = 10000;
	// 属性以及上面提示的乱七八糟的东西excel总行数 需要针对不同模板具体配置
	private static final int PropertyNum = 3;
	// 导入设备文件地址 可改动态配置
	private static final String filePath = "C:\\Users\\zhang\\Desktop\\xmlc\\out.xlsx";
	// 拆分文件夹存储路径 可改动态配置
	private static final String splitPath = "C:\\Users\\zhang\\Desktop\\xmlc\\splite\\";

	private int i = 1, day, month = 0, year = 0;
	private HashMap<String, CellStyle> datestyle = new HashMap();;
	private XSSFCellStyle cellStyles;

	public String splitExcels(String filePath, int FileMaxNum) {

		try {
			//清空分割存储文件夹
			File[] splitFiles = new File(splitPath).listFiles();
			;
			if (null != splitFiles) {
				for (File f2 : splitFiles) {

					f2.delete();
				}
				System.out.println("清理旧文件...");
			}

			File file = new File(filePath);
			System.out.println("等待拆分文件路径》》" + file.getAbsolutePath());
			Map<String, SXSSFWorkbook> map = getSplitMap(file.getAbsolutePath(), FileMaxNum);// 得到拆分后的子文件存储对象

			createSplitWorkbook(map, splitPath, file.getName());// 遍历对象生成的拆分文件
			// }
		} catch (Exception e) {
			e.printStackTrace();
		}
		return splitPath;
	}

	/**
	 * @Title: getSplitMap
	 * @Description:将第一列的值作为键值,将一个文件拆分为多个文件
	 * @param fileName
	 * @param FileMaxNum
	 * @return
	 * @throws Exception
	 * 
	 */
	public Map<String, SXSSFWorkbook> getSplitMap(String fileName, int FileMaxNum) throws Exception {
		System.out.println("开始拆分文件....");
		Map<String, SXSSFWorkbook> map = new HashMap<String, SXSSFWorkbook>();
		File fn = new File(fileName);
		InputStream is = new FileInputStream(fn);
		// 根据输入流创建Workbook对象
		XSSFWorkbook wb = getWorkbok(fn);
		// get到Sheet对象
		XSSFSheet sheet = wb.getSheetAt(0);
		// 获取文件行数
		int MaxNum = sheet.getLastRowNum();
		System.out.println("文件行行数" + MaxNum + "行.");
		// 文件限定行数
		if (FileMaxNum <= 0) {
			FileMaxNum = 10000;//
		}
		// FileMaxNum=10;//用于测试
		// 分割文件数量
		int fileNum = (MaxNum % FileMaxNum) == 0 ? MaxNum / FileMaxNum : (MaxNum / FileMaxNum) + 1;
		System.out.println("文件将要被拆分为" + fileNum + "个文件.");
		// Row titleRow = null;
		Row[] rows = new Row[PropertyNum];
		// 这个必须用接口
		int i = 0;
		// =null;
		SXSSFWorkbook tempWorkbook = null;
		SXSSFSheet secSheet = null;
		XSSFCellStyle cellStyle = null;
		for (Row row : sheet) {// 遍历每一行
			row = (XSSFRow) row;
			int currentFile = 1;
			int currentRowNum = row.getRowNum();
			for (int k = 0; k < fileNum; k++) {
				if ((currentRowNum >= k * FileMaxNum + rows.length)
						&& (currentRowNum < (k + 1) * FileMaxNum + rows.length)) {
					currentFile = k + 1;
					System.out.println("当前操作第" + currentFile + "个文件.");
				}

			}
			if (i < rows.length) {
				rows[i] = row;// 得到标题（加上些没用的）

			} else {
//			
				String key = String.valueOf(currentFile);
				tempWorkbook = map.get(key);
				if (tempWorkbook == null) {// 如果以当前行第一列值作为键值取不到工作表
					System.out.println("tempWorkbook为空 ，，新建");
					tempWorkbook = new SXSSFWorkbook(1);
					cellStyle = (XSSFCellStyle) tempWorkbook.createCellStyle();

					SXSSFSheet tempSheet = tempWorkbook.createSheet();
					secSheet = tempWorkbook.getSheetAt(0);
					for (int p = 0; p < rows.length; p++) {
						Row firstRow = tempSheet.createRow(p);
						for (short k = 0; k < rows[p].getLastCellNum(); k++) {// 为每个子文件创建标题
							Cell c = rows[p].getCell(k);
							// System.out.println(c);
							if (c != null) {
								// System.out.println(c);
								Cell newcell = firstRow.createCell(k);
								newcell.setCellValue(c.toString());
							}
						}
					}
					map.put(key, tempWorkbook);
				}

				SXSSFRow secRow = secSheet.createRow(secSheet.getLastRowNum() + 1);
				for (short m = 0; m < row.getLastCellNum(); m++) {
					SXSSFCell newcell = secRow.createCell(m);
					XSSFCell cell = (XSSFCell) row.getCell(m);

					setCellValue(newcell, cell, tempWorkbook);
				}
				map.put(key, tempWorkbook);
			}
			System.out.println("开始遍历行...." + i);
			i = i + 1;// 行数加一
		}
		return map;

	}

	/**
	 * @Title: createSplitWorkbook
	 * @Description:创建文件
	 * @param map
	 * @param savePath
	 * @param fileName
	 * @throws IOException
	 */
	public void createSplitWorkbook(Map<String, SXSSFWorkbook> map, String savePath, String fileName)
			throws IOException {

		Iterator iter = map.entrySet().iterator();
		while (iter.hasNext()) {
			Map.Entry entry = (Map.Entry) iter.next();
			String key = (String) entry.getKey();
			SXSSFWorkbook val = (SXSSFWorkbook) entry.getValue();
			File filePath = new File(savePath);
			if (!filePath.exists()) {
				// 存放目录不存在,自动为您创建存放目录
				filePath.mkdir();
			}
			if (!filePath.isDirectory()) {
				// 无效文件目录
				System.err.println("无效目录");
			}
			String filename = savePath + key + "_" + fileName;
			System.out.println("拆分后的新文件》》" + filename);
			File file = new File(filename);
			if (!file.exists()) {
				file.createNewFile();
			}
			FileOutputStream fOut;// 新建输出文件流
			try {
				fOut = new FileOutputStream(file);
				val.write(fOut); // 把相应的Excel工作薄存盘
				fOut.flush();
				fOut.close(); // 操作结束，关闭文件
				val.dispose();
				System.gc();
			} catch (FileNotFoundException e) {
				System.err.println("操作失败,如果文件已打开，请关闭excel文件后再试" + e.getMessage());
			} catch (Exception e) {
				System.err.println("操作失败" + e.getMessage());
			}
		}
		System.out.println("操作完成，文件位置：" + splitPath);

	}

	/**
	 * @Title: setCellValue
	 * @Description:将一个单元格的值赋给另一个单元格
	 * @param newCell
	 * @param cell
	 * @param wb
	 * @param cellStyle
	 * @throws InterruptedException
	 * @throws ParseException
	 */
	public void setCellValue(SXSSFCell newCell, XSSFCell cell, SXSSFWorkbook wb)
			throws InterruptedException, ParseException {

		if (cell == null) {
			return;
		}
		// newCell.setCellValue(cell.toString());
		switch (cell.getCellType()) {

		case BOOLEAN:
			newCell.setCellValue(cell.getBooleanCellValue());
			break;
		case NUMERIC:

			if (DateUtil.isCellDateFormatted(cell)) {
				
				//特殊处理 减少style对象创建  应对大量数据防止OOM
				String key = wb.hashCode() + cell.getCellStyle().getDataFormatString();
				cellStyles = (XSSFCellStyle) datestyle.get(key);
				if (cellStyles == null) {
					XSSFDataFormat format = (XSSFDataFormat) wb.createDataFormat();
					XSSFCellStyle cst = (XSSFCellStyle) wb.createCellStyle();
					cst.setDataFormat(format.getFormat(cell.getCellStyle().getDataFormatString()));
					cellStyles = cst;
					datestyle.put(key, cst);
					System.out.println("XSSFDataFormat创建>>>>" + datestyle.toString());
				}
				// cellStyles = (XSSFCellStyle) wb.createCellStyle();】
				//特殊格式特殊处理
				Date date = cell.getDateCellValue();
				if (i % 2 != 0) {

					day = date.getDate();
					month = date.getMonth();
					year = date.getYear();
//
				} else {
					date.setYear(year);
					date.setMonth(month);
					date.setDate(day);
					cell.setCellValue(date);

				}
				//System.out.println("datestyleformat..." + cellStyles.getDataFormatString());
				newCell.setCellStyle(cellStyles);
				newCell.setCellValue(cell.getDateCellValue());

				i = i + 1;

			} else {
				// 读取数字
				newCell.setCellValue(cell.getNumericCellValue());
			}
			break;
		case FORMULA:

			newCell.setCellValue(cell.getCellFormula());
			break;
		case STRING:
			newCell.setCellValue(cell.getStringCellValue());
			break;

		}
	}

	// 兼容 2003 2007+
	public XSSFWorkbook getWorkbok(File file) throws IOException {
		XSSFWorkbook wb = null;
		FileInputStream in = new FileInputStream(file);
		wb = new XSSFWorkbook(in);
		System.out.println("模板版本2007+excel");

		return wb;
	}

	public static void main(String[] arg) {
		SplitExcelUtil sp = new SplitExcelUtil();
		new Thread(new Runnable() {

			@Override
			public void run() {
				// TODO Auto-generated method stub
				String splitPath = sp.splitExcels(filePath, FileMaxNum);
			}

		}).start();

	}
}
