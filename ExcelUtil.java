package com.zzl.util;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.text.NumberFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 *@author zzl
 *@since 2018.11.19
 */
public class ExcelUtil {
    /***
     * @param sheets
     * @param fields
     * @return map集合
     */
    public static List<Map<String,Object>> getFilds(HSSFWorkbook sheets, List<String> fields){
        HSSFSheet sheet = sheets.getSheetAt(0);
        List<Map<String,Object>> keyValueList = new ArrayList<Map<String, Object>>();
        for (int i = 1;i<=sheet.getLastRowNum();i++) {
            Map<String,Object> map = new HashMap<String, Object>();
            Row row = sheet.getRow(i);
            for (int j = 0;j<fields.size();j++) {
                NumberFormat nf = NumberFormat.getInstance();
                map.put(fields.get(j),getValue(row.getCell(j)));
            }
            keyValueList.add(map);
        }
        return keyValueList;
    }

    private static Object getValue(Cell cell) {
        Object obj = null;
        NumberFormat nf = NumberFormat.getInstance();
        switch (cell.getCellType()) {
            case BOOLEAN:
                obj = cell.getBooleanCellValue();
                break;
            case ERROR:
                obj = cell.getErrorCellValue();
                break;
            case NUMERIC:
                obj = nf.format(cell.getNumericCellValue());
                break;
            case STRING:
                obj = cell.getStringCellValue();
                break;
            default:
                break;
        }
        return obj;
    }
}
