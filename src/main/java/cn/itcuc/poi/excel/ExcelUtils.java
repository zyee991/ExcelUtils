package cn.itcuc.poi.excel;

import java.beans.BeanInfo;
import java.beans.Introspector;
import java.beans.PropertyDescriptor;
import java.io.InputStream;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelUtils {
	public static List<Map<String,Object>> importExcel(InputStream in) throws Exception {
		List<Map<String,Object>> mapList = new ArrayList<Map<String,Object>>();
		Workbook workbook = WorkbookFactory.create(in);
		Sheet sheet = workbook.getSheetAt(0);
		Row row = sheet.getRow(0);
		List<String> keys = new ArrayList<String>();
		for(int i = 0; i < row.getLastCellNum(); i++){
			Cell cell = row.getCell(i);
			keys.add(String.valueOf(getValue(cell)));
		}
		
		for(int i = 0; i < sheet.getLastRowNum(); i++){
			Row currentRow = sheet.getRow(i + 1);
			Map<String, Object> map = new HashMap<String, Object>();
			for(int j = 0; j < currentRow.getLastCellNum(); j++){
				map.put(keys.get(j), getValue(currentRow.getCell(j)));
			}
			mapList.add(map);
		}
		return mapList;
	}
	
	public static <T> List<T> importExcel(InputStream in, Class<T> c) throws Exception {
		List<T> list = new ArrayList<T>();
		Workbook workbook = WorkbookFactory.create(in);
		Sheet sheet = workbook.getSheetAt(0);
		Row row = sheet.getRow(0);
		List<String> keys = new ArrayList<String>();
		for(int i = 0; i < row.getLastCellNum(); i++){
			Cell cell = row.getCell(i);
			keys.add(String.valueOf(getValue(cell)));
		}
		
		for(int i = 0; i < sheet.getLastRowNum(); i++){
			Row currentRow = sheet.getRow(i + 1);
			Map<String, Object> map = new HashMap<String, Object>();
			for(int j = 0; j < currentRow.getLastCellNum(); j++){
				map.put(keys.get(j), getValue(currentRow.getCell(j)));
			}
			T t = mapToObject(c,map);
			list.add(t);
		}
		
		return list;
	}
	
	private static <T> T mapToObject(Class<T> c,Map<String, Object> map) throws Exception {
		BeanInfo beanInfo = Introspector.getBeanInfo(c);
		T t = c.newInstance();
		PropertyDescriptor[] propertyDescriptors = beanInfo.getPropertyDescriptors();
		for(int i = 0; i < propertyDescriptors.length; i++){
			PropertyDescriptor descriptor = propertyDescriptors[i];
			String propertyName = descriptor.getName();
			if(map.containsKey(propertyName)){
				Object value = map.get(propertyName);
				Object[] args = new Object[1];
				args[0] = value;
				try {
					descriptor.getWriteMethod().invoke(t, args);
				} catch (IllegalAccessException e) {
					e.printStackTrace();
				} catch (IllegalArgumentException e) {
					e.printStackTrace();
				} catch (InvocationTargetException e) {
					e.printStackTrace();
				}
			}
		}
		return t;
	}
	
	private static Object getValue(Cell cell) {
        if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
            return cell.getBooleanCellValue();
        } else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
            return cell.getNumericCellValue();
        } else {
            return String.valueOf(cell.getStringCellValue());
        }
    }
}
