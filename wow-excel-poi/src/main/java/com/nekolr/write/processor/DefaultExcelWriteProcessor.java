package com.nekolr.write.processor;

import com.nekolr.exception.ExcelWriteInitException;
import com.nekolr.metadata.Excel;
import com.nekolr.metadata.ExcelField;
import com.nekolr.metadata.Head;
import com.nekolr.util.BeanUtils;
import com.nekolr.util.ExcelUtils;
import com.nekolr.util.ParameterUtils;
import com.nekolr.write.ExcelWriteContext;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.Comparator;
import java.util.List;
import java.util.Optional;


public class DefaultExcelWriteProcessor implements ExcelWriteProcessor {

    private ExcelWriteContext writeContext;

    @Override
    public void init(ExcelWriteContext writeContext) {
        this.writeContext = writeContext;
        this.createWorkbook();
    }

    @Override
    public void write(List<?> data) {
        this.writeHead();
        if (data != null) {
            Excel excel = this.writeContext.getExcel();
            int rowIndex = this.writeContext.getRowIndex();
            int colIndex = this.writeContext.getColIndex();
            Sheet sheet = this.getOrCreateSheet();
            List<ExcelField> excelFieldList = excel.getFieldList();
            for (int i = 0; i < data.size(); i++) {
                Object entity = data.get(i);
                Row row = sheet.createRow(rowIndex);
                for (int col = 0; col < excelFieldList.size(); col++) {
                    ExcelField excelField = excelFieldList.get(col);
                    Cell cell = row.createCell(colIndex + col);
                    // 获取属性值
                    Object attrValue = BeanUtils.getFieldValue(entity, excelField.getField());
                    // 单元格赋值
                    ExcelUtils.setCellValue(cell, attrValue, excelField);
                }
            }
        }
    }

    protected void writeHead() {
        Sheet sheet = this.getOrCreateSheet();
        Excel excel = this.writeContext.getExcel();
        int rowIndex = this.writeContext.getRowIndex();
        int colIndex = this.writeContext.getColIndex();
        List<ExcelField> excelFieldList = excel.getFieldList();
        // 上一个列的索引
        int lastColIndex = colIndex;
        // 遍历拿到多级表头的最大层数
        Optional<Integer> max = excelFieldList.stream()
                .map(field -> field.getFiledName().split(excel.getTitleSeparator()).length)
                .filter(level -> level > 1)
                .max(Comparator.naturalOrder());
        if (max.isPresent()) { // 如果有多级表头
            // 最大合并行数
            Integer maxRowSpan = max.get();
            Row row = sheet.createRow(rowIndex);
            // 将字段元数据列表转换成 Head 类型的列表
            List<Head> headList = ExcelUtils.toTree(ExcelUtils.toHeadList(excel));
            // 遍历转换后的表头列表
            for (int multiCol = 0; multiCol < headList.size(); multiCol++) {
                Head head = headList.get(multiCol);
                List<Head> children = head.getChildren();
                Cell cell = row.createCell(colIndex + multiCol);
                // 设置表头单元格内容
                ExcelUtils.setCellValue(cell, head.getTitle(), head.getExcelField());
                if (children.size() == 0) { // 单级表头
                    // 需要合并最大的行数
                    sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex + maxRowSpan - 1, lastColIndex, lastColIndex));
                    lastColIndex++;
                } else if (head.getLevel() < max.get()) { // 多级表头，但不是层数最深的那个
                    // 需要合并的列数，为当前表头的所有叶子表头的个数
                    int colspan = ExcelUtils.getLeafChildCount(head, 0);
                    // 需要合并的行数
                    int rowspan = maxRowSpan - head.getLevel() + 1;
                    int firstRow = rowIndex;
                    int lastRow = rowIndex + rowspan - 1;
                    int firstCol = lastColIndex;
                    int lastCol = lastColIndex + colspan - 1;
                    sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));

                    drillDown(sheet, children, lastRow, firstCol, 0);
                    lastColIndex += colspan;
                } else { // 多级表头，最深的那个
                    int rowspan = 1;
                    int colspan = ExcelUtils.getLeafChildCount(head, 0);
                    int firstRow = rowIndex;
                    int lastRow = rowIndex + rowspan - 1;
                    int firstCol = lastColIndex;
                    int lastCol = lastColIndex + colspan - 1;
                    sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));

                    drillDown(sheet, children, lastRow, firstCol, 0);
                    lastColIndex += colspan;
                }
            }
        } else { // 没有多级表头

        }
    }

    private void drillDown(Sheet sheet, List<Head> child, int lastRow, int firstCol, int lastCol) {
        for (int i = 0; i < child.size(); i++) {
            int rowspan = 1;
            int colspan = ExcelUtils.getLeafChildCount(child.get(i), 0);
            if (colspan != 1) {
                // TODO: 此处有问题
                if (i == 0) {
                    lastCol = firstCol + colspan - 1;
                    sheet.addMergedRegion(new CellRangeAddress(lastRow + 1, lastRow + rowspan, firstCol, lastCol));
                } else {
                    sheet.addMergedRegion(new CellRangeAddress(lastRow + 1, lastRow + rowspan, lastCol + 1, lastCol + colspan));
                }
                // lastCol 在下钻时只是作为一个初始值即可
                drillDown(sheet, child.get(i).getChildren(), lastRow + 1, firstCol, 0);
            }

        }
    }

    protected void writeBigTitle() {

    }

    /**
     * 创建 Workbook
     *
     * @return Workbook
     */
    private void createWorkbook() {
        Workbook workbook = this.writeContext.getWorkbook();
        if (workbook == null) {
            // TODO: 如果没有设置 Excel 注解，这里就是空
            Excel excel = this.writeContext.getExcel();
            switch (excel.getWorkbookType()) {
                case XLS:
                    this.writeContext.setWorkbook(new HSSFWorkbook());
                    break;
                case XLSX:
                    if (this.writeContext.isStreamingWriterEnabled()) {
                        this.writeContext.setWorkbook(new SXSSFWorkbook());
                    } else {
                        this.writeContext.setWorkbook(new XSSFWorkbook());
                    }
                    break;
                default:
                    throw new ExcelWriteInitException("Unsupported format: " + excel.getWorkbookType().name());
            }
        }
    }

    /**
     * 创建 Sheet
     *
     * @return Sheet
     */
    private Sheet getOrCreateSheet() {
        Workbook workbook = this.writeContext.getWorkbook();
        String sheetName = this.writeContext.getSheetName();
        Sheet sheet = this.writeContext.getSheet();
        if (sheet == null) {
            if (ParameterUtils.isEmpty(sheetName)) {
                sheet = workbook.createSheet();
            } else {
                sheet = workbook.createSheet(sheetName);
            }
            this.writeContext.setSheet(sheet);
        }
        return sheet;
    }
}
