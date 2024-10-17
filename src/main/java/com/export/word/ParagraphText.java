package com.export.word;

import org.docx4j.wml.R;
import org.docx4j.wml.Text;

import java.util.List;

/**
 * 段落的最小部分，由Run和Text组成
 */
public class ParagraphText extends Element {

    private String value;
    private ParagraphTextStyle paragraphTextStyle;

    public ParagraphText() {

    }

    public ParagraphText(String value) {
        this.value = value;
        paragraphTextStyle = ParagraphTextStyle.DEFAULT;
    }

    public ParagraphText(String value, Font font, FontSize fontSize) {
        this.value = value;
        this.paragraphTextStyle = new ParagraphTextStyle(font, fontSize);
    }

    public ParagraphText(String value, Font font, FontSize fontSize, Boolean fontBold) {
        this.value = value;
        this.paragraphTextStyle = new ParagraphTextStyle(font, fontSize, fontBold);
    }

    @Override
    public void toElement(WordExportContext ctx, List<Object> content) {
        R r = new R();
        content.add(r);
        // 设置样式
        if (this.paragraphTextStyle != null) {
            r.setRPr(this.paragraphTextStyle.get());
        }
        // 添加text
        Text text = new Text();
        r.getContent().add(text);
        // 设置值
        text.setValue(this.value);
    }

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }

    public ParagraphTextStyle getParagraphTextStyle() {
        return paragraphTextStyle;
    }

    public void setParagraphTextStyle(ParagraphTextStyle paragraphTextStyle) {
        this.paragraphTextStyle = paragraphTextStyle;
    }

}
