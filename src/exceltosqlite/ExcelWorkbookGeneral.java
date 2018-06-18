package exceltosqlite;

import java.io.File;
import java.io.IOException;

import javax.swing.JFileChooser;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * 获取excel文件中的数据
 * 
 * @author shetia
 * @version v1.0
 */
public class ExcelWorkbookGeneral {

	String ExcelWorkbook_path = null;
	String ExcelWorkbook_sheetname = null;
	Workbook ExcelWorkbook_rawdata = null;
	Sheet ExcelWorkbook_sheet_rawdata = null;
	Row ExcelWorkbook_row_rawdata = null;
	Cell ExcelWorkbook_cell_rawdata = null;
	String ExcelWorkbook_cellvalue = null;

	int ExcelWorkbook_rowNumbertotal = -1;
	int anchor_row = -1;
	int anchor_cell = -1;

	/**
	 * 使用文件路径打开Excel文件
	 * @param ExcelWorkbook_path excel文件的绝对路径
	 * @return wbFile 返回已打开的Workbook类型对象
	 */
	public static Workbook OpenExeclWorkbook(String ExcelWorkbook_path)
			throws EncryptedDocumentException, InvalidFormatException, IOException {
		// 使用传入参数打开excel文件
		File f = new File(ExcelWorkbook_path);
		System.out.println(f.getParent() + " " + f.getName());
		Workbook wbFile = WorkbookFactory.create(f);
		return wbFile;
	}

	/**
	 * 使用可视化界面打开Excel文件
	 * @return wbFile 返回已打开的Workbook类型对象
	 */
	public static Workbook OpenExeclWorkbook()
			throws EncryptedDocumentException, InvalidFormatException, IOException {
		//使用gui打开excel文件
		File f = FileChooser();
		String filename = f.getName();
		String filedir = f.getParent();
		System.out.println(filedir + " " + filename);
		Workbook wbFile = WorkbookFactory.create(f);		
		return wbFile;
	}
	
	/**
	 * 使用swing选择需要打开的excel文件,返回File对象
	 * @return file 返回file类型对象
	 */
	private static File FileChooser() {
		JFileChooser fc = new JFileChooser("C:");
		// 是否可多选
		fc.setMultiSelectionEnabled(false);
		// 选择模式，可选择文件和文件夹
		// fc.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);
		fc.setFileSelectionMode(JFileChooser.FILES_ONLY);
		// fc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
		// 设置是否显示隐藏文件
		fc.setFileHidingEnabled(true);
		fc.setAcceptAllFileFilterUsed(false);
		// 设置文件筛选器
		// fc.setFileFilter(new MyFilter("java"));
		fc.setFileFilter(new FileNameExtensionFilter("Excel文件(*.xls|*.xlsx)", "xls", "xlsx"));
		int returnValue = fc.showOpenDialog(null);
		if (returnValue == JFileChooser.APPROVE_OPTION) {
			File file = fc.getSelectedFile();
			file = new File(file.getAbsolutePath());
			return file;
		}
		return null;
	}
	
	/**
	 * 关闭已打开Excel文件
	 * @param ExcelWorkbook_path excel文件的绝对路径
	 * @return 无
	 */
	public void CloseExeclWorkbook(Workbook wbFile) throws IOException {
		// 关闭文件
		wbFile.close();
	}
	
	/**
	 * 根据已打开的Workbook对象和对应的sheet名,打开指定名字的sheet对象
	 * @param wbFile 已打开的Workbook对象
	 * @param ExcelWorkbook_sheetname 表名sheetname
	 * @return sheet_data sheetname对应的sheet对象
	 */
	public Sheet LocateExcel_Sheet(Workbook wbFile, String ExcelWorkbook_sheetname) {
		// 查找excel表中特定名字的sheet对象
		for (Sheet sheet_data : wbFile) {
			if (sheet_data.getSheetName().equals(ExcelWorkbook_sheetname)) {
				return sheet_data;
			}
		}
		return null;
	}

	/**
	 * 返回特定sheet对象中存在的实际单元格行数(存在实际行数多余非空行数的情况,原因是之前存在数据然后被清除掉)
	 * @param sheet_data 已打开的sheet对象
	 * @return Sheet_PhysicalNumberOfRows 返回初始化过的行数(Row Number)
	 */
	public int GetSheet_RowNumber(Sheet sheet_data) {
		int Sheet_PhysicalNumberOfRows = sheet_data.getPhysicalNumberOfRows();
		return Sheet_PhysicalNumberOfRows;
	}
	
	/**
	 * 根据sheet对象和row定位锚点获取行对象
	 * @param anchor_row row定位锚点
	 * @param sheet_data 已打开的sheet对象
	 * @return row_data row定位锚点所在行(row)对象
	 */
	public Row GetSheet_RowData(int anchor_row, Sheet sheet_data) {
		// 获取特定行数据(raw)
		Row row_data = sheet_data.getRow(anchor_row);
		return row_data;
	}
	
	/**
	 * 根据row对象和cell定位锚点获取单元格对象
	 * 对未初始化的单元格置空
	 * @param anchor_cell cell定位锚点
	 * @param row_data 已打开的row对象
	 * @return cell_data cell定位锚点所在的单元格(cell)对象
	 */
	public Cell GetSheet_CellData(int anchor_cell, Row row_data) {
		// 获取特定行特定单元格数据(raw)
		Cell cell_data = row_data.getCell(anchor_cell,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
		return cell_data;
	}
	
	/**
	 * 根据单元格对象获取单元格中的值
	 * 如果单元格不是string类型则设置为String类型
	 * @param cell_data 已打开的单元格(cell)对象
	 * @return cell_data_value 单元格的值
	 */
	public String GetSheet_CellDataValue(Cell cell_data) {
		cell_data.setCellType(CellType.STRING);
		String cell_data_value = cell_data.getStringCellValue();
		return cell_data_value;
	}
}
