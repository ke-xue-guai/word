package com.export.word;

import org.docx4j.wml.BooleanDefaultTrue;
import org.docx4j.wml.PPr;

/**
 * 段落样式
 */
public class ParagraphStyle extends Single<PPr> {

    private Spacing spacing;
    private TextAlign textAlign;
    private boolean pageBreakBefore = false;

    @Override
    public PPr create() {
        PPr ppr = new PPr();
        if (this.spacing != null) {
            ppr.setSpacing(this.spacing.get());
        }
        if (this.textAlign != null) {
            ppr.setJc(this.textAlign.get());
        }
        BooleanDefaultTrue bo = new BooleanDefaultTrue();
        bo.setVal(this.pageBreakBefore);
        ppr.setPageBreakBefore(bo);
        return ppr;
    }

    public Spacing getSpacing() {
        return spacing;
    }

    public void setSpacing(Spacing spacing) {
        this.spacing = spacing;
    }

    public TextAlign getTextAlign() {
        return textAlign;
    }

    public void setTextAlign(TextAlign textAlign) {
        this.textAlign = textAlign;
    }

    public void setPageBreakBefore(boolean pageBreakBefore) {
        this.pageBreakBefore = pageBreakBefore;
    }

}
