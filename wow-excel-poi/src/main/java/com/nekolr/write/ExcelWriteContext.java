package com.nekolr.write;

import com.nekolr.Constants;
import com.nekolr.metadata.Excel;
import lombok.Getter;
import lombok.Setter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.OutputStream;

@Getter
@Setter
public class ExcelWriteContext {

    /**
     * 输出文件名称
     * <p>
     * 在没有指定输出文件的情况下，优先使用该设置，其次使用 Excel 注解上的 value
     */
    private String filename;

    /**
     * 输出的 Excel 文件
     */
    private File file;

    /**
     * 输出流
     */
    private OutputStream outputStream;

    /**
     * Excel 注解的元数据
     *
     * @see com.nekolr.annotation.Excel
     */
    private Excel excel;

    /**
     * 忽略的字段
     */
    private String[] ignores;

    /**
     * sheet 名称
     */
    private String sheetName;

    /**
     * 文档密码
     */
    private String password;

    /**
     * workbook
     */
    private Workbook workbook;

    /**
     * sheet
     */
    private Sheet sheet;

    /**
     * 是否开启流式写，默认关闭
     */
    private boolean streamingWriterEnabled = Constants.STREAMING_WRITER_DISABLED;

    /**
     * 开始的行号
     */
    private int rowIndex = Constants.DEFAULT_WRITE_ROW_INDEX;

    /**
     * 开始的列号
     */
    private int colIndex = Constants.DEFAULT_WRITE_COL_INDEX;

}
