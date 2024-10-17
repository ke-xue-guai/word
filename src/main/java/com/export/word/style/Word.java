package com.export.word.style;

import lombok.Data;

import java.util.List;


@Data
public class Word {

    /**
     * 表格行
     */
    private int rows;

    /**
     * 表格列
     */
    private int cols;

    /**
     * 单元格
     */
    private List<Cell> cellList;

    /**
     * 重复行
     */
    private int repeatNum;

    /**
     * 重复列
     */
    private int repeatColNum;

    /**
     * 纸张初始样式
     */
    private PageStyle pageStyle;

    /**
     * 表格整体样式
     */
    private TableStyle tableStyle;

    /**
     * 页眉
     */
    private HeadStyle headStyle;
}
