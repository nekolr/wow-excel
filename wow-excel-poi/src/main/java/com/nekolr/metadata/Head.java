package com.nekolr.metadata;

import lombok.Getter;
import lombok.Setter;

import java.util.List;
import java.util.Objects;

/**
 * 带有上下级关系的表头
 */
@Getter
@Setter
public class Head {

    /**
     * 表头标题
     */
    private String title;

    /**
     * 表头在字段元数据列表中的索引
     */
    private Integer oldIndex;

    /**
     * 表头深度
     */
    private Integer level;

    /**
     * 表头对应的字段元数据
     */
    private ExcelField excelField;

    /**
     * 父标题
     */
    private String parentTitle;

    /**
     * 子表头列表
     */
    private List<Head> children;

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        Head head = (Head) o;
        return title.equals(head.title);
    }

    @Override
    public int hashCode() {
        return Objects.hash(title);
    }
}
