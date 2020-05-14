package com.nekolr.write.merge;

import lombok.Data;

/**
 * 上一个单元格的信息
 */
@Data
public class LastCell {

    /**
     * 上一个单元格的值
     */
    private Object value;

    /**
     * 上一个单元格的列号
     */
    private int colNum;
}
