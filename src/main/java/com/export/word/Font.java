package com.export.word;

import org.docx4j.wml.RFonts;
import org.docx4j.wml.STHint;

public class Font extends Single<RFonts> {

    public static final Font DEFAULT = new Font("微软雅黑");

    private String name;

    public Font(String name) {
        this.name = name;
    }

    @Override
    public RFonts create() {
        // 设置文本 用于指定ASCII,HAnsi,复合字体字符集的字体名称
        RFonts font = new RFonts();
        font.setCs(this.name);
        font.setAscii(this.name);
        font.setHAnsi(this.name);
        font.setEastAsia(this.name);
        font.setHint(STHint.EAST_ASIA);
        return font;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

}
