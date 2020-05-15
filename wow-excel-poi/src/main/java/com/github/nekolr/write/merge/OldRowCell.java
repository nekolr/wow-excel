package com.github.nekolr.write.merge;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * 旧的行所在列的单元格的信息
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
public class OldRowCell {

    /**
     * 旧的行所在列的单元格的值
     */
    private Object cellValue;

    /**
     * 旧的行号
     */
    private int rowNum;

}
