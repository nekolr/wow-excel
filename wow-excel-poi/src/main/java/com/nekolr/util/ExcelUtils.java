package com.nekolr.util;

import com.nekolr.metadata.DataConverter;
import com.nekolr.metadata.ExcelField;
import com.nekolr.write.merge.LastCell;
import com.nekolr.write.merge.OldRowCell;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

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
                return ParamUtils.numeric2FieldType(cell.getNumericCellValue(), field);
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
     * 使用写转换器
     *
     * @param attrValue     属性数据
     * @param field         字段元数据
     * @param dataConverter 使用的数据转换器
     * @return 转换后的数据
     */
    public static Object useWriteConverter(Object attrValue, ExcelField field, DataConverter dataConverter) {
        return dataConverter.toExcelAttribute(field, attrValue);
    }

    /**
     * 合并列
     *
     * @param lastCell    上一个单元格
     * @param sheet       sheet
     * @param row         行
     * @param firstColNum 表格的起始列号
     * @param colNum      列号
     * @param colSize     列数
     * @param cellValue   单元格值
     */
    public static void mergeCols(LastCell lastCell, Sheet sheet, Row row, int firstColNum, int colNum, int colSize, Object cellValue) {
        // 将第一列的单元格信息保存
        if (colNum == firstColNum) {
            lastCell.setValue(cellValue);
            lastCell.setColNum(colNum);
            return;
        }
        // 上一个单元格与当前单元格的值相同，说明需要合并列
        if (ParamUtils.equals(cellValue, lastCell.getValue())) {
            // 直到当前列为最后一列时，合并列
            if (colNum - firstColNum == colSize - 1) {
                sheet.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum(), lastCell.getColNum(), colNum));
            }
            return;
        }
        // 值相同的单元格有时不会连续出现直到最后一列
        // 此时 lastCell 记录的还是第一次出现的单元格，并且当前单元格的值一定是第一个和记录值不同的单元格（因为相同的值会走上面的逻辑）
        // 此时 lastCell 记录的单元格的列号 + 1 还是小于当前列号，说明多次出现了值相同的列但是没有进行合并
        if (lastCell.getColNum() + 1 < colNum) {
            sheet.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum(), lastCell.getColNum(), colNum - 1));
        }
        // 将每一列的信息保存，最后一列不需要保存，因为它没有下一列，即没有某个列的上一列是它
        // 如果遇到需要合并的列（即上一个列的值和该列的值相同），lastCell 保存的是需要合并的第一个单元格
        if (colNum - firstColNum != colSize - 1) {
            lastCell.setValue(cellValue);
            lastCell.setColNum(colNum);
        }

    }

    /**
     * 合并行
     *
     * @param oldRowCache 旧的行数据缓存
     * @param sheet       sheet
     * @param row         当前行
     * @param rowIndex    遍历行时的索引
     * @param colNum      列号
     * @param rowSize     行数
     * @param cellValue   单元格数据
     */
    public static void mergeRows(Map<Integer, OldRowCell> oldRowCache, Sheet sheet, Row row, int rowIndex,
                                 int colNum, int rowSize, Object cellValue) {
        // 将第一行每个表头单元格的信息记录下来
        if (rowIndex == 0) {
            oldRowCache.put(colNum, new OldRowCell(cellValue, row.getRowNum()));
            return;
        }
        OldRowCell oldRowCell = oldRowCache.get(colNum);
        // 值相同说明需要合并行
        if (ParamUtils.equals(cellValue, oldRowCell.getCellValue())) {
            // 直到当前行为最后一行时，进行合并行
            if (rowIndex == rowSize - 1) {
                sheet.addMergedRegion(new CellRangeAddress(oldRowCell.getRowNum(), row.getRowNum(), colNum, colNum));
            }
            return;
        }
        // 与合并列类似，值相同的单元格也有可能不会连续出现直到最后一行，这里就是对这种情况进行补漏地合并
        if (oldRowCell.getRowNum() + 1 < row.getRowNum()) {
            sheet.addMergedRegion(new CellRangeAddress(oldRowCell.getRowNum(), row.getRowNum() - 1, colNum, colNum));
        }
        // 执行到这里说明值不相等，意味着不需要合并行，所以更新行所在列的信息，保存上一行所在列的单元格信息
        if (rowIndex != rowSize - 1) {
            oldRowCache.put(colNum, new OldRowCell(cellValue, row.getRowNum()));
        }
    }
}
