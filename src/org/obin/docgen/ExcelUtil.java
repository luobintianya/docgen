/**
 *
 */
package org.obin.docgen;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @author wade
 */
public class ExcelUtil {

	private final Workbook workbook;
	private FileOutputStream outputStream;
	private String pattern;// 日期格式
	private XSSFFormulaEvaluator xsEvaluator;// 日期格式

	public void setPattern(final String pattern) {
		this.pattern = pattern;
	}

	public ExcelUtil(String fileoutput) throws Exception {
		this.outputStream = new FileOutputStream(fileoutput);

		workbook = new XSSFWorkbook();
	}

	public ExcelUtil(final File file) throws IOException, Exception {
		workbook = new XSSFWorkbook(file);
		if (workbook instanceof XSSFWorkbook) {
			xsEvaluator = new XSSFFormulaEvaluator((XSSFWorkbook) workbook);
		}
	}

	public ExcelUtil(final InputStream is) throws IOException {
		workbook = new XSSFWorkbook(is);
		if (workbook instanceof XSSFWorkbook) {
			xsEvaluator = new XSSFFormulaEvaluator((XSSFWorkbook) workbook);
		}
	}

	public ExcelUtil(final Workbook workboook) {
		this.workbook = workboook;
		if (workbook instanceof XSSFWorkbook) {
			xsEvaluator = new XSSFFormulaEvaluator((XSSFWorkbook) workbook);
		}
	}

	public void writeExcel(List<List<String>> list) throws FileNotFoundException {

		// 创建工作�? 

		// 创建工作�?
		XSSFSheet xssfSheet;
		xssfSheet = (XSSFSheet) workbook.createSheet();

		// 创建�?
		XSSFRow xssfRow;

		// 创建列，即单元格Cell
		XSSFCell xssfCell;

		// 把List里面的数据写到excel�?
		for (int i = 0; i < list.size(); i++) {
			// 从第�?行开始写�?
			xssfRow = xssfSheet.createRow(i);
			// 创建每个单元格Cell，即列的数据
			List sub_list = list.get(i);
			for (int j = 0; j < sub_list.size(); j++) {
				xssfCell = xssfRow.createCell(j); // 创建单元�?
				xssfCell.setCellValue((String) sub_list.get(j)); // 设置单元格内�?
			}
		}

		// 用输出流写到excel
		try {
			workbook.write(outputStream);
			outputStream.flush();
			outputStream.close();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	@Override
	public String toString() {

		return "共有 " + getSheetCount() + "个sheet 页！";
	}

	public String toString(final int sheetIx) throws IOException {

		return "�? " + (sheetIx + 1) + "个sheet 页，名称�?  " + getSheetName(sheetIx) + "，共 " + getRowCount(sheetIx) + "行！";
	}

	/**
	 * 根据后缀判断是否�? Excel 文件，后�?匹配xls和xlsx
	 *
	 * @param pathname
	 * @return
	 */
	public static boolean isExcel(final String pathname) {
		if (pathname == null) {
			return false;
		}
		return pathname.endsWith(".xls") || pathname.endsWith(".xlsx");
	}

	/**
	 * 读取 Excel 第一页所有数�?
	 *
	 * @return
	 * @throws Exception
	 */
	public List<Map<String, String>> read() throws Exception {
		return read(0, 0, getRowCount(0) - 1);
	}

	public List<Map<String, String>> read(int sheet, int rowNumber) throws Exception {
		return read(0, rowNumber, getRowCount(0) - 1);
	}

	public List<Map<String, String>> read(int sheet, int rowNumber, boolean dateissue) throws Exception {
		return read(0, rowNumber, getRowCount(0) - 1, dateissue);
	}

	/**
	 * 读取指定sheet 页所有数�?
	 *
	 * @param sheetIx
	 *            指定 sheet 页，�? 0 �?�?
	 * @return
	 * @throws Exception
	 */
	public List<Map<String, String>> read(final int sheetIx) throws Exception {
		return read(sheetIx, 0, getRowCount(sheetIx) - 1);
	}

	/**
	 * 读取指定sheet 页指定行数据
	 *
	 * @param sheetIx
	 *            指定 sheet 页，�? 0 �?�?
	 * @param start
	 *            指定�?始行，从 0 �?�?
	 * @param end
	 *            指定结束行，�? 0 �?�?
	 * @return
	 * @throws Exception
	 */
	public List<Map<String, String>> read(final int sheetIx, final int start, int end) throws Exception {

		final Sheet sheet = workbook.getSheetAt(sheetIx);
		final List<Map<String, String>> list = new LinkedList<>();

		if (end > getRowCount(sheetIx)) {
			end = getRowCount(sheetIx);
		}

		final Row attrRow = sheet.getRow(0);
		final int cols = attrRow.getLastCellNum(); // 第一行�?�列�?

		for (int i = start; i <= end; i++) {
			final Map<String, String> rowList = new HashMap();
			final Row row = sheet.getRow(i);
			if (row == null) {
				continue;
			}
			for (int j = 0; j < cols; j++) {
				if (null != getCellValueToString(row.getCell(j))) {
					rowList.put(getCellValueToString(attrRow.getCell(j)), getCellValueToString(row.getCell(j)).trim());
				} else {
					continue;
				}
			}

			if (rowList.size() > 0)
				list.add(rowList);
		}

		return list;
	}

	public List<Map<String, String>> read(final int sheetIx, final int start, int end, boolean dateissue)
			throws Exception {
		final Sheet sheet = workbook.getSheetAt(sheetIx);
		final List<Map<String, String>> list = new LinkedList<>();

		if (end > getRowCount(sheetIx)) {
			end = getRowCount(sheetIx);
		}

		final Row attrRow = sheet.getRow(0);
		final int cols = attrRow.getLastCellNum(); // 第一行�?�列�?

		for (int i = start; i <= end; i++) {
			final Map<String, String> rowList = new HashMap();
			final Row row = sheet.getRow(i);
			if (row == null) {
				continue;
			}
			for (int j = 0; j < cols; j++) {
				if (null != getCellValueToString(row.getCell(j), dateissue)) {
					rowList.put(getCellValueToString(attrRow.getCell(j)),
							getCellValueToString(row.getCell(j), dateissue).trim());
				} else {
					continue;
				}
			}

			if (rowList.size() > 0)
				list.add(rowList);
		}

		return list;
	}

	/**
	 * 将数据写入到 Excel 默认第一页中，从�?1行开始写�?
	 *
	 * @param rowData
	 *            数据
	 * @return
	 * @throws IOException
	 */
	public boolean write(final List<List<String>> rowData) throws IOException {
		return write(0, rowData, 0);
	}

	/**
	 * 将数据写入到 Excel 新创建的 Sheet �?
	 *
	 * @param rowData
	 *            数据
	 * @param sheetName
	 *            长度�?1-31，不能包含后面任�?字符: ：\ / ? * [ ]
	 * @return
	 * @throws IOException
	 */
	public boolean write(final List<List<String>> rowData, final String sheetName, final boolean isNewSheet)
			throws IOException {
		Sheet sheet = null;
		if (isNewSheet) {
			sheet = workbook.createSheet(sheetName);
		} else {
			sheet = workbook.createSheet();
		}
		final int sheetIx = workbook.getSheetIndex(sheet);
		return write(sheetIx, rowData, 0);
	}

	/**
	 * 将数据追加到sheet页最�?
	 *
	 * @param rowData
	 *            数据
	 * @param sheetIx
	 *            指定 Sheet 页，�? 0 �?�?
	 * @param isAppend
	 *            是否追加,true 追加，false 重置sheet再添�?
	 * @return
	 * @throws IOException
	 */
	public boolean write(final int sheetIx, final List<List<String>> rowData, final boolean isAppend)
			throws IOException {
		if (isAppend) {
			return write(sheetIx, rowData, getRowCount(sheetIx));
		} else {// 清空再添�?
			clearSheet(sheetIx);
			return write(sheetIx, rowData, 0);
		}
	}

	/**
	 * 将数据写入到 Excel 指定 Sheet 页指定开始行�?,指定行后面数据向后移�?
	 *
	 * @param rowData
	 *            数据
	 * @param sheetIx
	 *            指定 Sheet 页，�? 0 �?�?
	 * @param startRow
	 *            指定�?始行，从 0 �?�?
	 * @return
	 * @throws IOException
	 */
	public boolean write(final int sheetIx, final List<List<String>> rowData, final int startRow) throws IOException {
		final Sheet sheet = workbook.getSheetAt(sheetIx);
		final int dataSize = rowData.size();
		if (getRowCount(sheetIx) > 0) {// 如果小于等于0，则�?行都不存�?
			sheet.shiftRows(startRow, getRowCount(sheetIx), dataSize);
		}
		for (int i = 0; i < dataSize; i++) {
			final Row row = sheet.createRow(i + startRow);
			for (int j = 0; j < rowData.get(i).size(); j++) {
				final Cell cell = row.createCell(j);
				cell.setCellValue(rowData.get(i).get(j) + "");
			}
		}
		return true;
	}

	/**
	 * 设置cell 样式
	 *
	 * @param sheetIx
	 *            指定 Sheet 页，�? 0 �?�?
	 * @param colIndex
	 *            指定列，�? 0 �?�?
	 * @return
	 * @throws IOException
	 */
	public boolean setStyle(final int sheetIx, final int rowIndex, final int colIndex, final CellStyle style)
			throws IOException {
		final Sheet sheet = workbook.getSheetAt(sheetIx);
		// sheet.autoSizeColumn(colIndex, true);// 设置列宽度自适应
		sheet.setColumnWidth(colIndex, 4000);

		final Cell cell = sheet.getRow(rowIndex).getCell(colIndex);
		cell.setCellStyle(style);

		return true;
	}

	/**
	 * 设置样式
	 *
	 * @param type
	 *            1：标�? 2：第�?�?
	 * @return
	 */
	public CellStyle makeStyle(final int type) {
		final CellStyle style = workbook.createCellStyle();

		final DataFormat format = workbook.createDataFormat();
		style.setDataFormat(format.getFormat("@"));// // 内容样式 设置单元格内容格式是文本
		// style.setAlignment(CellStyle.ALIGN_CENTER);// 内容居中

		// style.setBorderTop(CellStyle.BORDER_THIN);// 边框样式
		// style.setBorderRight(CellStyle.BORDER_THIN);
		// style.setBorderBottom(CellStyle.BORDER_THIN);
		// style.setBorderLeft(CellStyle.BORDER_THIN);

		final Font font = workbook.createFont();// 文字样式

		if (type == 1) {
			// style.setFillForegroundColor(HSSFColor.LIGHT_BLUE.index);//颜色样式
			// 前景颜色
			// style.setFillBackgroundColor(HSSFColor.LIGHT_BLUE.index);//背景�?
			// style.setFillPattern(CellStyle.ALIGN_FILL);// 填充方式
			font.setBold(true);
			font.setFontHeight((short) 500);
		}

		if (type == 2) {
			font.setBold(true);
			font.setFontHeight((short) 300);
		}

		style.setFont(font);

		return style;
	}

	/**
	 * 合并单元�?
	 *
	 * @param sheetIx
	 *            指定 Sheet 页，�? 0 �?�?
	 * @param firstRow
	 *            �?始行
	 * @param lastRow
	 *            结束�?
	 * @param firstCol
	 *            �?始列
	 * @param lastCol
	 *            结束�?
	 */
	public void region(final int sheetIx, final int firstRow, final int lastRow, final int firstCol,
			final int lastCol) {
		final Sheet sheet = workbook.getSheetAt(sheetIx);
		sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
	}

	/**
	 * 指定行是否为�?
	 *
	 * @param sheetIx
	 *            指定 Sheet 页，�? 0 �?�?
	 * @param rowIndex
	 *            指定�?始行，从 0 �?�?
	 * @return true 不为空，false 不行为空
	 * @throws IOException
	 */
	public boolean isRowNull(final int sheetIx, final int rowIndex) throws IOException {
		final Sheet sheet = workbook.getSheetAt(sheetIx);
		return sheet.getRow(rowIndex) == null;
	}

	/**
	 * 创建行，若行存在，则清空
	 *
	 * @param sheetIx
	 *            指定 sheet 页，�? 0 �?�? 指定创建行，�? 0 �?�?
	 * @return
	 * @throws IOException
	 */
	public boolean createRow(final int sheetIx, final int rowIndex) throws IOException {
		final Sheet sheet = workbook.getSheetAt(sheetIx);
		sheet.createRow(rowIndex);
		return true;
	}

	/**
	 * 指定单元格是否为�?
	 *
	 * @param sheetIx
	 *            指定 Sheet 页，�? 0 �?�?
	 * @param rowIndex
	 *            指定�?始行，从 0 �?�?
	 * @param colIndex
	 *            指定�?始列，从 0 �?�?
	 * @return true 行不为空，false 行为�?
	 * @throws IOException
	 */
	public boolean isCellNull(final int sheetIx, final int rowIndex, final int colIndex) throws IOException {
		final Sheet sheet = workbook.getSheetAt(sheetIx);
		if (!isRowNull(sheetIx, rowIndex)) {
			return false;
		}
		final Row row = sheet.getRow(rowIndex);
		return row.getCell(colIndex) == null;
	}

	/**
	 * 创建单元�?
	 *
	 * @param sheetIx
	 *            指定 sheet 页，�? 0 �?�?
	 * @param rowIndex
	 *            指定行，�? 0 �?�?
	 * @param colIndex
	 *            指定创建列，�? 0 �?�?
	 * @return true 列为空，false 行不为空
	 * @throws IOException
	 */
	public boolean createCell(final int sheetIx, final int rowIndex, final int colIndex) throws IOException {
		final Sheet sheet = workbook.getSheetAt(sheetIx);
		final Row row = sheet.getRow(rowIndex);
		row.createCell(colIndex);
		return true;
	}

	/**
	 * 返回sheet 中的行数
	 *
	 * @param sheetIx
	 *            指定 Sheet 页，�? 0 �?�?
	 * @return
	 */
	public int getRowCount(final int sheetIx) {
		final Sheet sheet = workbook.getSheetAt(sheetIx);
		if (sheet.getPhysicalNumberOfRows() == 0) {
			return 0;
		}
		return sheet.getLastRowNum() + 1;

	}

	/**
	 * 返回�?在行的列�?
	 *
	 * @param sheetIx
	 *            指定 Sheet 页，�? 0 �?�?
	 * @param rowIndex
	 *            指定行，�?0�?�?
	 * @return 返回-1 表示�?在行为空
	 */
	public int getColumnCount(final int sheetIx, final int rowIndex) {
		final Sheet sheet = workbook.getSheetAt(sheetIx);
		final Row row = sheet.getRow(rowIndex);
		return row == null ? -1 : row.getLastCellNum();

	}

	/**
	 * 设置row �? column 位置的单元格�?
	 *
	 * @param sheetIx
	 *            指定 Sheet 页，�? 0 �?�?
	 * @param rowIndex
	 *            指定行，�?0�?�?
	 * @param colIndex
	 *            指定列，�?0�?�?
	 * @param value
	 *            �?
	 * @return
	 * @throws IOException
	 */
	public boolean setValueAt(final int sheetIx, final int rowIndex, final int colIndex, final String value)
			throws IOException {
		final Sheet sheet = workbook.getSheetAt(sheetIx);
		sheet.getRow(rowIndex).getCell(colIndex).setCellValue(value);
		return true;
	}

	/**
	 * 返回 row �? column 位置的单元格�?
	 *
	 * @param sheetIx
	 *            指定 Sheet 页，�? 0 �?�?
	 * @param rowIndex
	 *            指定行，�?0�?�?
	 * @param colIndex
	 *            指定列，�?0�?�?
	 * @return
	 */
	public String getValueAt(final int sheetIx, final int rowIndex, final int colIndex) {
		final Sheet sheet = workbook.getSheetAt(sheetIx);
		return getCellValueToString(sheet.getRow(rowIndex).getCell(colIndex));
	}

	/**
	 * 重置指定行的�?
	 *
	 * @param rowData
	 *            数据
	 * @param sheetIx
	 *            指定 Sheet 页，�? 0 �?�?
	 * @param rowIndex
	 *            指定行，�?0�?�?
	 * @return
	 * @throws IOException
	 */
	public boolean setRowValue(final int sheetIx, final List<String> rowData, final int rowIndex) throws IOException {
		final Sheet sheet = workbook.getSheetAt(sheetIx);
		final Row row = sheet.getRow(rowIndex);
		for (int i = 0; i < rowData.size(); i++) {
			row.getCell(i).setCellValue(rowData.get(i));
		}
		return true;
	}

	/**
	 * 返回指定行的值的集合
	 *
	 * @param sheetIx
	 *            指定 Sheet 页，�? 0 �?�?
	 * @param rowIndex
	 *            指定行，�?0�?�?
	 * @return
	 */
	public List<String> getRowValue(final int sheetIx, final int rowIndex) {
		final Sheet sheet = workbook.getSheetAt(sheetIx);
		final Row row = sheet.getRow(rowIndex);
		final List<String> list = new ArrayList<String>();
		if (row == null) {
			list.add(null);
		} else {
			for (int i = 0; i < row.getLastCellNum(); i++) {
				String value = getCellValueToString(row.getCell(i));
				if (value != null) {
					list.add(value);
				}
			}
		}
		return list;
	}

	/**
	 * 返回列的值的集合
	 *
	 * @param sheetIx
	 *            指定 Sheet 页，�? 0 �?�?
	 * @param rowIndex
	 *            指定行，�?0�?�?
	 * @param colIndex
	 *            指定列，�?0�?�?
	 * @return
	 */
	public List<String> getColumnValue(final int sheetIx, final int rowIndex, final int colIndex) {
		final Sheet sheet = workbook.getSheetAt(sheetIx);
		final List<String> list = new ArrayList<String>();
		for (int i = rowIndex; i < getRowCount(sheetIx); i++) {
			final Row row = sheet.getRow(i);
			if (row == null) {
				list.add(null);
				continue;
			}

			String value = getCellValueToString(sheet.getRow(i).getCell(colIndex));
			if (value != null) {
				list.add(value);
			}

		}
		return list;
	}

	/**
	 * 获取excel 中sheet 总页�?
	 *
	 * @return
	 */
	public int getSheetCount() {
		return workbook.getNumberOfSheets();
	}

	public void createSheet() {
		workbook.createSheet();
	}

	/**
	 * 设置sheet名称，长度为1-31，不能包含后面任�?字符: ：\ / ? * [ ]
	 *
	 * @param sheetIx
	 *            指定 Sheet 页，�? 0 �?始，//
	 * @param name
	 * @return
	 * @throws IOException
	 */
	public boolean setSheetName(final int sheetIx, final String name) throws IOException {
		workbook.setSheetName(sheetIx, name);
		return true;
	}

	/**
	 * 获取 sheet名称
	 *
	 * @param sheetIx
	 *            指定 Sheet 页，�? 0 �?�?
	 * @return
	 * @throws IOException
	 */
	public String getSheetName(final int sheetIx) throws IOException {
		final Sheet sheet = workbook.getSheetAt(sheetIx);
		return sheet.getSheetName();
	}

	/**
	 * 获取sheet的索引，�?0�?�?
	 *
	 * @param name
	 *            sheet 名称
	 * @return -1表示该未找到名称对应的sheet
	 */
	public int getSheetIndex(final String name) {
		return workbook.getSheetIndex(name);
	}

	/**
	 * 删除指定sheet
	 *
	 * @param sheetIx
	 *            指定 Sheet 页，�? 0 �?�?
	 * @return
	 * @throws IOException
	 */
	public boolean removeSheetAt(final int sheetIx) throws IOException {
		workbook.removeSheetAt(sheetIx);
		return true;
	}

	/**
	 * 删除指定sheet中行，改变该行之后行的索�?
	 *
	 * @param sheetIx
	 *            指定 Sheet 页，�? 0 �?�?
	 * @param rowIndex
	 *            指定行，�?0�?�?
	 * @return
	 * @throws IOException
	 */
	public boolean removeRow(final int sheetIx, final int rowIndex) throws IOException {
		final Sheet sheet = workbook.getSheetAt(sheetIx);
		sheet.shiftRows(rowIndex + 1, getRowCount(sheetIx), -1);
		final Row row = sheet.getRow(getRowCount(sheetIx) - 1);
		sheet.removeRow(row);
		return true;
	}

	/**
	 * 设置sheet 页的索引
	 *
	 * @param sheetname
	 *            Sheet 名称 Sheet 索引，从0�?�?
	 */
	public void setSheetOrder(final String sheetname, final int sheetIx) {
		workbook.setSheetOrder(sheetname, sheetIx);
	}

	/**
	 * 清空指定sheet页（先删除后添加并指定sheetIx�?
	 *
	 * @param sheetIx
	 *            指定 Sheet 页，�? 0 �?�?
	 * @return
	 * @throws IOException
	 */
	public boolean clearSheet(final int sheetIx) throws IOException {
		final String sheetname = getSheetName(sheetIx);
		removeSheetAt(sheetIx);
		workbook.createSheet(sheetname);
		setSheetOrder(sheetname, sheetIx);
		return true;
	}

	public Workbook getWorkbook() {
		return workbook;
	}

	/**
	 * 关闭�?
	 *
	 * @throws IOException
	 */
	public void close() throws IOException {
		if (outputStream != null) {
			outputStream.close();
		}
		workbook.close();
	}

	/**
	 * 转换单元格的类型为String 默认�? <br>
	 * 默认的数据类型：CELL_TYPE_BLANK(3), CELL_TYPE_BOOLEAN(4),
	 * CELL_TYPE_ERROR(5),CELL_TYPE_FORMULA(2), CELL_TYPE_NUMERIC(0),
	 * CELL_TYPE_STRING(1)
	 *
	 * @param cell
	 * @return
	 */
	private String getCellValueToString(final Cell cell) {
		String strCell = null;
		if (cell == null) {
			return strCell;
		}
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_BOOLEAN:
			strCell = String.valueOf(cell.getBooleanCellValue());
			break;
		case Cell.CELL_TYPE_NUMERIC:
			if (HSSFDateUtil.isCellDateFormatted(cell)) {
				final Date date = cell.getDateCellValue();
				if (pattern != null) {
					final SimpleDateFormat sdf = new SimpleDateFormat(pattern);
					strCell = sdf.format(date);
				} else {
					strCell = date.toString();
				}
				break;
			}
			// 不是日期格式，则防止当数字过长时以科学计数法显示
			cell.setCellType(HSSFCell.CELL_TYPE_STRING);
			strCell = cell.toString();
			break;
		case Cell.CELL_TYPE_STRING:
			strCell = cell.getStringCellValue();
			break;

		case Cell.CELL_TYPE_FORMULA:
			if (xsEvaluator != null) {
				CellValue CellValue = xsEvaluator.evaluate(cell);
				return String.valueOf(CellValue.formatAsString());
			}
			break;
		case Cell.CELL_TYPE_BLANK:
			strCell = null;
			break;
		default:
			break;
		}
		if (strCell != null && strCell.trim().length() == 0) {
			strCell = null;
		}

		return strCell;
	}

	private String getCellValueToString(final Cell cell, boolean tmp) {
		String strCell = null;
		if (cell == null) {
			return strCell;
		}
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_BOOLEAN:
			strCell = String.valueOf(cell.getBooleanCellValue());
			break;
		case Cell.CELL_TYPE_NUMERIC:
			if (HSSFDateUtil.isCellDateFormatted(cell)) {
				final Date date = cell.getDateCellValue();
				if (pattern != null) {
					final SimpleDateFormat sdf = new SimpleDateFormat(pattern);
					strCell = sdf.format(date);
				} else {
					strCell = date.toString();
				}
				break;
			}
			// 不是日期格式，则防止当数字过长时以科学计数法显示
			if (!tmp) {
				cell.setCellType(HSSFCell.CELL_TYPE_STRING);
			}
			strCell = cell.toString();
			break;
		case Cell.CELL_TYPE_STRING:
			strCell = cell.getStringCellValue();
			break;
		case Cell.CELL_TYPE_BLANK:
			strCell = null;
			break;
		default:
			break;
		}
		if (strCell != null && strCell.trim().length() == 0) {
			strCell = null;
		}

		return strCell;
	}
}