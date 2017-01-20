package cn.itcuc.poi.excel;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;

public class ExcelUtils {
	public static <T> T upload(InputStream in,Class<T> c) throws IOException, InstantiationException, IllegalAccessException {
		POIFSFileSystem fileSystem = new POIFSFileSystem(in);
		HSSFWorkbook workbook = new HSSFWorkbook(fileSystem);
		HSSFSheet sheet = workbook.getSheetAt(0);
		HSSFRow row = sheet.getRow(0);
		List<String> keys = new ArrayList<String>();
		for(int i = 0; i < row.getLastCellNum(); i++){
			HSSFCell cell = row.getCell(i);
			keys.add(getValue(cell));
		}
		
		T t = c.newInstance();
		List<String> filedList = new ArrayList<String>();
		Field[] fileds = c.getFields();
		for(Field filed : fileds){
			filedList.add(filed.getName());
		}
		
		return null;
	}
	
	 @SuppressWarnings("deprecation")
	private static String getValue(HSSFCell hssfCell) {
	        if (hssfCell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
	            // 返回布尔类型的值
	            return String.valueOf(hssfCell.getBooleanCellValue());
	        } else if (hssfCell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
	            // 返回数值类型的值
	            return String.valueOf(hssfCell.getNumericCellValue());
	        } else {
	            // 返回字符串类型的值
	            return String.valueOf(hssfCell.getStringCellValue());
	        }
	    }
}
