package com.export.word;

import org.docx4j.wml.BooleanDefaultTrue;
import org.docx4j.wml.CTVerticalAlignRun;
import org.docx4j.wml.RPr;
import org.docx4j.wml.STVerticalAlignRun;

public class ParagraphTextStyle extends Single<RPr> {

    public static final ParagraphTextStyle DEFAULT = new ParagraphTextStyle(Font.DEFAULT, FontSize.DEFAULT);

    private Font font;
    private FontSize fontSize;
    private FontColor fontColor;
    private Boolean fontBold;

    public ParagraphTextStyle(Font font, FontSize fontSize) {
        this.font = font;
        this.fontSize = fontSize;
    }

    public ParagraphTextStyle(Font font, FontSize fontSize, Boolean fontBold) {
        this.font = font;
        this.fontSize = fontSize;
        this.fontBold = fontBold;
    }

    @Override
    public RPr create() {
        RPr rpr = new RPr();
        // 默认对齐方式
        CTVerticalAlignRun verticalAlign = new CTVerticalAlignRun();
        verticalAlign.setVal(STVerticalAlignRun.BASELINE);
        rpr.setVertAlign(verticalAlign);
        // 设置字体
        if (this.font != null) {
            rpr.setRFonts(this.font.get());
        }
        // 设置字号
        if (this.fontSize != null) {
            rpr.setSz(this.fontSize.get());
        }
        // 字体颜色
        if (this.fontColor != null) {
            rpr.setColor(this.fontColor.get());
        }
        // 设置加粗
        if (this.fontBold != null) {
            BooleanDefaultTrue booleanDefaultTrue = new BooleanDefaultTrue();
            booleanDefaultTrue.setVal(this.fontBold);
            rpr.setB(booleanDefaultTrue);
        }
        return rpr;
    }

    public Font getFont() {
        return font;
    }

    public void setFont(Font font) {
        this.font = font;
    }

    public FontSize getFontSize() {
        return fontSize;
    }

    public void setFontSize(FontSize fontSize) {
        this.fontSize = fontSize;
    }

    public FontColor getFontColor() {
        return fontColor;
    }

    public void setFontColor(FontColor fontColor) {
        this.fontColor = fontColor;
    }

}
