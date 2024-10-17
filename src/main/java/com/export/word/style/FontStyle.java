package com.export.word.style;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class FontStyle {

    /**
     * 默认字体样式
     */
    public final static FontStyle DEFAULT_FONT_STYLE = new FontStyle();

    /**
     * 是否加粗
     */
    private boolean bold;

    /**
     * 字体
     */
    private String fontName;

    /**
     * 字体大小
     */
    private int fontSize;

    /**
     * 字体颜色
     */
    private String wordColor;

    /**
     * 是否添加着重号
     */
    private boolean em;

    /**
     * 是否倾斜
     */
    private boolean lean;

    /**
     * 是否有下划线
     */
    private boolean underType;

    /**
     * 下划线颜色
     */
    private String underColor;
}
