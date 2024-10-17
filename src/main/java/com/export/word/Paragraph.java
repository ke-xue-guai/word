package com.export.word;

import org.docx4j.wml.P;
import org.docx4j.wml.PPr;
import org.docx4j.wml.PPrBase;
import org.docx4j.wml.STLineSpacingRule;

import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;

public class Paragraph extends Element {

    private ParagraphStyle paragraphStyle;
    private List<ParagraphText> texts = new ArrayList<>();

    public Paragraph() {
    }

    public Paragraph(String value) {
        this.texts.add(new ParagraphText(value));
    }

    public Paragraph(String value, Font font, FontSize fontSize, TextAlign textAlign) {
        ParagraphText text = new ParagraphText(value, font, fontSize);
        this.texts.add(text);
        if (textAlign != null) {
            this.paragraphStyle = new ParagraphStyle();
            this.paragraphStyle.setTextAlign(textAlign);
        }
    }

    @Override
    public void toElement(WordExportContext ctx, List<Object> content) {
        P p = new P();
        content.add(p);
        // 添加文本
        for (ParagraphText text : texts) {
            text.toElement(ctx, p.getContent());
        }
        // 添加段落样式
        if (this.paragraphStyle != null) {
            p.setPPr(this.paragraphStyle.get());
//            if (ctx.isFixRowHigh) {
////                setLineHeight(p, ctx.rowHigh.divide(BigInteger.valueOf(3)));
//                setLineHeight(p, ctx.rowHigh);
//            }
        }
    }

    /**
     * 设置段落的行高
     */
    private static void setLineHeight(P paragraph, BigInteger rowHigh) {
        PPr pPr = paragraph.getPPr();
        PPrBase.Spacing spacing = pPr.getSpacing();
        if (spacing == null) {
            spacing = new PPrBase.Spacing();
            pPr.setSpacing(spacing);
        }
        // 设置行高
        spacing.setLine(rowHigh);
        // 设置行高类型为固定值
        spacing.setLineRule(STLineSpacingRule.EXACT);
    }

    public void addText(ParagraphText text) {
        this.texts.add(text);
    }

    public ParagraphStyle getParagraphStyle() {
        return paragraphStyle;
    }

    public void setParagraphStyle(ParagraphStyle paragraphStyle) {
        this.paragraphStyle = paragraphStyle;
    }

    public List<ParagraphText> getTexts() {
        return texts;
    }

    public void setTexts(List<ParagraphText> texts) {
        this.texts = texts;
    }

}
