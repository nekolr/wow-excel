package com.nekolr.util;

import com.nekolr.metadata.ExcelFieldBean;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;

import java.lang.reflect.Field;

public class ExcelUtils {

    /**
     * 获取单元格的值
     *
     * @param cell           单元格
     * @param excelFieldBean 字段元数据的实体
     * @param field          对应的字段
     * @return 单元格的值
     */
    public static Object getCellValue(Cell cell, ExcelFieldBean excelFieldBean, Field field) {
        switch (cell.getCellType()) {
            case _NONE:
            case BLANK:
            case ERROR:
                return null;
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue();
                }
                return ParameterUtils.numeric2FieldType(cell.getNumericCellValue(), field);
            case FORMULA:
                return cell.getStringCellValue();
            default:
                // STRING
                if (excelFieldBean.isAutoTrim()) {
                    return cell.getStringCellValue().trim();
                } else {
                    return cell.getStringCellValue();
                }
        }
    }

    public static Object useConverter(Object value) {
        return value;
    }

}
