package com.nekolr.read.listener;

import com.nekolr.metadata.ExcelField;
import com.nekolr.metadata.ExcelReadResultListener;
import com.nekolr.read.ExcelReadContext;

import java.util.List;

/**
 * Excel 读事件处理器
 */
public class ExcelReadEventProcessor {

    /**
     * 在读 sheet 之前
     *
     * @param readListeners 读监听器集合
     * @param readContext   读上下文
     * @param <R>           使用 @Excel 注解的类类型
     */
    public static <R> void beforeReadSheet(List<ExcelSheetReadListener<R>> readListeners, ExcelReadContext<R> readContext) {
        if (readListeners != null) {
            readListeners.forEach(listener -> listener.beforeReadSheet(readContext));
        }
    }

    /**
     * 在读一个单元格之后
     *
     * @param readListeners 读监听器集合
     * @param field         字段元数据
     * @param cellValue     单元格值
     * @param rowNum        行号
     * @param colNum        列号
     * @return 处理后的单元格值
     */
    public static Object afterReadCell(List<ExcelCellReadListener> readListeners, ExcelField field,
                                       Object cellValue, int rowNum, int colNum) {
        if (readListeners != null) {
            for (ExcelCellReadListener readListener : readListeners) {
                cellValue = readListener.afterReadCell(field, cellValue, rowNum, colNum);
            }
        }
        return cellValue;
    }

    /**
     * 在读一个空白单元格之后
     *
     * @param readListeners 读监听器集合
     * @param field         单元格对应的表头字段的元数据
     * @param rowNum        所在行
     * @param colNum        所在列
     * @return 处理后的单元格值
     */
    public static Object afterReadEmptyCell(List<ExcelEmptyCellReadListener> readListeners, ExcelField field,
                                            int rowNum, int colNum) {
        Object cellValue = null;
        if (readListeners != null) {
            for (ExcelEmptyCellReadListener readListener : readListeners) {
                cellValue = readListener.handleEmpty(field, rowNum, colNum);
            }
        }
        return cellValue;
    }

    /**
     * 在读一行之后
     *
     * @param readListeners 读监听器集合
     * @param r             一行的数据，对应一个实体对象
     * @param rowNum        行号
     * @param <R>           使用 @Excel 注解的类类型
     * @return 是否停止读
     */
    public static <R> boolean afterReadRow(List<ExcelRowReadListener<R>> readListeners, R r, int rowNum) {
        boolean isStop = false;
        if (readListeners != null) {
            for (ExcelRowReadListener<R> readListener : readListeners) {
                isStop = readListener.afterReadRow(r, rowNum);
            }
        }
        return isStop;
    }

    /**
     * 在读完 sheet 之后
     *
     * @param readListeners 读监听器集合
     * @param readContext   Excel 读上下文
     * @param <R>           使用 @Excel 注解的类类型
     */
    public static <R> void afterReadSheet(List<ExcelSheetReadListener<R>> readListeners, ExcelReadContext<R> readContext) {
        if (readListeners != null) {
            readListeners.forEach(listener -> listener.afterReadSheet(readContext));
        }
    }

    /**
     * 结果通知
     *
     * @param resultListener 结果通知监听器
     * @param result         读取结果
     * @param <R>            使用 @Excel 注解的类类型
     */
    public static <R> void resultNotify(ExcelReadResultListener<R> resultListener, List<R> result) {
        if (result != null && resultListener != null) {
            resultListener.notify(result);
        }
    }
}
