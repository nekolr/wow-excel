package com.nekolr.util;

import com.nekolr.metadata.DataConverter;
import com.nekolr.metadata.Excel;
import com.nekolr.metadata.ExcelField;
import com.nekolr.metadata.Head;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;

import java.lang.reflect.Field;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;

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
     * 设置单元格的值
     *
     * @param cell  单元格
     * @param value 值
     */
    public static void setCellValue(Cell cell, Object value, ExcelField field) {
        if (value == null) {
            // do nothing
        } else if (value instanceof String) {
            cell.setCellValue(value.toString());
        } else if (value instanceof Number) {
            cell.setCellValue(((Number) value).doubleValue());
        } else if (value instanceof Date) {
            cell.setCellValue((Date) value);
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        } else if (value instanceof LocalDate) {
            cell.setCellValue((LocalDate) value);
        } else if (value instanceof LocalDateTime) {
            cell.setCellValue((LocalDateTime) value);
        } else if (value instanceof Enum) {
            cell.setCellValue(value.toString());
        } else {
            throw new IllegalArgumentException("Unsupported data type, field: " + field.getField().getName() + " value: " + value);
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

    /**
     * 将列表转换成树形列表
     * <p>
     * 时间复杂度：O(N)
     *
     * @param list 未形成树形结构的列表
     * @return 形成树形结构的列表
     */
    public static List<Head> toTree(List<Head> list) {
        if (list != null) {
            Map<String, List<Head>> map = new HashMap<>((int) (list.size() / 0.75F + 1.0F));
            List<Head> result = new ArrayList<>();
            list.forEach(head -> {
                map.computeIfAbsent(head.getTitle(), k -> new ArrayList<>());
                head.setChildren(map.get(head.getTitle()));
                if (head.getParentTitle() == null) {
                    result.add(head);
                } else {
                    map.computeIfAbsent(head.getParentTitle(), k -> new ArrayList<>());
                    map.get(head.getParentTitle()).add(head);
                    map.put(head.getParentTitle(), map.get(head.getParentTitle()));
                }
            });
            return result;
        }
        return null;
    }

    /**
     * 将表头字段元数据列表转换成 Head 类型的列表
     *
     * @param excel excel 注解元数据
     * @return Head 类型的列表
     */
    public static List<Head> toHeadList(Excel excel) {
        List<ExcelField> excelFieldList = excel.getFieldList();
        Map<String, Head> map = new LinkedHashMap<>();
        for (int i = 0; i < excelFieldList.size(); i++) {
            ExcelField field = excelFieldList.get(i);
            // 使用表头标题分隔符拆分
            String[] titles = field.getFiledName().split(excel.getTitleSeparator());
            for (int j = 0; j < titles.length; j++) {
                Head head = new Head();
                head.setTitle(titles[j]);
                // 设置层数
                head.setLevel(titles.length - j);
                if (j > 0) {
                    // 后边的才有父标题
                    head.setParentTitle(titles[j - 1]);
                }
                if (j == titles.length - 1) {
                    // 最后一个标题才有字段元数据
                    head.setExcelField(field);
                    head.setOldIndex(i);
                }
                if (map.containsKey(head.getTitle())) {
                    // 如果已有的元素的 level 比新元素的小，则进行覆盖
                    if (head.getLevel() > map.get(head.getTitle()).getLevel()) {
                        map.put(head.getTitle(), head);
                    }
                } else {
                    map.put(head.getTitle(), head);
                }
            }
        }
        return new ArrayList(map.values());
    }

    /**
     * 计算给定表头的叶子表头个数
     *
     * @param head  指定的表头
     * @param count 叶子表头个数
     * @return 叶子表头个数
     */
    public static int getLeafChildCount(Head head, int count) {
        List<Head> children = head.getChildren();
        if (children.size() == 0) {
            count++;
        } else {
            for (Head child : children) {
                count = getLeafChildCount(child, count);
            }
        }
        return count;
    }
}
