package com.export.word;

import org.docx4j.wml.Color;

/**
 * 字体颜色
 */
public class FontColor extends Single<Color> {

    private String value = "000000";

    @Override
    public Color create() {
        Color color = new Color();
        color.setVal(value);
        return color;
    }

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }

}
