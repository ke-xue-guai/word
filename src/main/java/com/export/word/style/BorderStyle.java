package com.export.word.style;

import lombok.Getter;
import lombok.Setter;


@Setter
@Getter
public class BorderStyle{

    /**
     * 边框样式：无、有边框、加粗边框
     */
    public enum Type {
        NIL, NONE, THICK, SINGLE, DOUBLE
    }

    /**
     * 左边框色
     */
    private String leftBorderColor;

    /**
     * 右边框色
     */
    private String rightBorderColor;

    /**
     * 上边框色
     */
    private String topBorderColor;

    /**
     * 下边框色
     */
    private String bottomBorderColor;

    /**
     * 左边框样式（是否隐藏）
     */
    private Type leftType;

    /**
     * 右边框样式
     */
    private Type rightType;

    /**
     * 上边框样式
     */
    private Type topType;

    /**
     * 下边框样式
     */
    private Type bottomType;

    /**
     * 边框大小
     */
    private Integer borderSize;

}
