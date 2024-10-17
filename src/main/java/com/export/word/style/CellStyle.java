package com.export.word.style;

import lombok.Getter;
import lombok.Setter;


@Setter
@Getter
public class CellStyle {

    public static final CellStyle DEFAULT = new CellStyle();
    /**
     * 单元格内垂直文本位置
     */
    public enum VertAlign {
        VERT_ALIGN_TOP, VERT_ALIGN_CENTER, VERT_ALIGN_BOTTOM
    }

    /**
     * 单元格内垂直水平位置
     */
    public enum Alignment {
        ALIGNMENT_LEFT, ALIGNMENT_CENTER, ALIGNMENT_RIGHT
    }

    /**
     * 单元格左边距
     */
    private float left;

    /**
     * 单元格右边距
     */
    private float right;

    /**
     * 单元格上边边距
     */
    private float top;

    /**
     * 单元格下边距
     */
    private float bottom;

    /**
     * 单元格背景色
     */
    private String bgColor;

    /**
     * 单元格内垂直水平位置
     */
    private VertAlign vertAlign;

    /**
     * 单元格内垂直文本位置
     */
    private Alignment alignment;

}
