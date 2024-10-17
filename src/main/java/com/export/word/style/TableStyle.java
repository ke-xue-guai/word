package com.export.word.style;

import lombok.Getter;
import lombok.Setter;

import java.util.List;
import java.util.Map;


@Getter
@Setter
public class TableStyle{

    /**
     * 行数
     */
    private int rowNum;

    /**
     * 内容的总高度
     */
    private Integer contentHeight;

    /**
     * 指定行高
     */
    private Map<Integer, Integer> rowHeight;

    /**
     * 表格宽度
     */
    private Integer tableWidth;

    /**
     * 最长的列宽格式
     */
    private List<WidthStyle> widthStyles;
}
