package com.nekolr.util;

import com.nekolr.metadata.DataConverter;
import com.nekolr.metadata.ExcelField;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;

import java.lang.reflect.Field;

public class ExcelUtils {

    /**
     * 获取单元格的值
     *
     * @param cell       单元格
     * @param excelField 字段元数据的实体
     * @param field      对应的字段
     * @return 单元格的值
     */
    public static Object getCellValue(Cell cell, ExcelField excelField, Field field) {
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
                if (excelField.isAutoTrim()) {
                    return cell.getStringCellValue().trim();
                } else {
                    return cell.getStringCellValue();
                }
        }
    }

    /**
     * 使用读转换器
     *
     * @param value         单元格数据
     * @param field         字段元数据
     * @param dataConverter 使用的数据转换器
     * @return 转换后的数据
     */
    public static Object useReadConverter(Object value, ExcelField field, DataConverter dataConverter) {
        return dataConverter.toEntityAttribute(field, value);
    }

}
