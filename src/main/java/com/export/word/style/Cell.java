package com.export.word.style;

import lombok.Data;

@Data
public class Cell {

    /**
     * 行
     */
    private int positionX;

    /**
     * 列
     */
    private int positionY;

    /**
     * 单元格所对应的值
     */
    private String value;

    /**
     * 合并列
     */
    private int colspan;

    /**
     * 合并行
     */
    private int rowspan;

    /**
     * 段落、文本、单元格样式
     */
    private BorderStyle borderStyle;

    /**
     * 单元格
     */
    private CellStyle cellStyle;

    /**
     * 字体
     */
    private FontStyle fontStyle;

    /**
     * 段落
     */
    private ParagraphStyle paragraphStyle;

    /**
     * 列宽
     */
    private WidthStyle widthStyle;
}
