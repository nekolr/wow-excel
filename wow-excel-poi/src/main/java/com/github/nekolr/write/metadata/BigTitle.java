package com.github.nekolr.write.metadata;

import lombok.*;

/**
 * 大标题
 */
@Getter
@Setter
@AllArgsConstructor
@NoArgsConstructor
@Builder
public final class BigTitle {

    /**
     * 行数
     */
    @Builder.Default
    private int lines = 2;

    /**
     * 起始列号
     */
    @Builder.Default
    private int firstColNum = 0;

    /**
     * 结束列号
     */
    @Builder.Default
    private int lastColNum = 0;

    /**
     * 内容
     */
    @Builder.Default
    private String content = "";
}
