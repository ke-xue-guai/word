package com.export.word;

import org.docx4j.wml.PPrBase;

import java.math.BigInteger;

/**
 * 段落的间距
 */
public class Spacing extends Single<PPrBase.Spacing> {

    private Long after;
    private Long before;
    private Long afterLines;
    private Long beforeLines;

    public Spacing() {

    }

    public Spacing(Long one) {
        this.after = one;
        this.before = one;
        this.afterLines = one;
        this.beforeLines = one;
    }

    public Spacing(Long one, Long two) {
        this.after = one;
        this.before = one;
        this.afterLines = two;
        this.beforeLines = two;
    }

    public Spacing(Long after, Long before, Long afterLines, Long beforeLines) {
        this.after = after;
        this.before = before;
        this.afterLines = afterLines;
        this.beforeLines = beforeLines;
    }


    @Override
    public PPrBase.Spacing create() {
        PPrBase.Spacing spacing = new PPrBase.Spacing();
        // 文本之后
        if (this.afterLines != null) {
            spacing.setAfterLines(BigInteger.valueOf(this.afterLines));
        }
        // 文本之前
        if (this.beforeLines != null) {
            spacing.setBeforeLines(BigInteger.valueOf(this.beforeLines));
        }
        // 段后距
        if (this.after != null) {
            spacing.setAfter(BigInteger.valueOf(this.after));
            spacing.setAfterAutospacing(false);
        }
        // 段前距
        if (this.before != null) {
            spacing.setBefore(BigInteger.valueOf(this.before));
            spacing.setBeforeAutospacing(false);
        }
        return spacing;
    }

    public long getAfter() {
        return after;
    }

    public void setAfter(long after) {
        this.after = after;
    }

    public long getBefore() {
        return before;
    }

    public void setBefore(long before) {
        this.before = before;
    }

    public long getAfterLines() {
        return afterLines;
    }

    public void setAfterLines(long afterLines) {
        this.afterLines = afterLines;
    }

    public long getBeforeLines() {
        return beforeLines;
    }

    public void setBeforeLines(long beforeLines) {
        this.beforeLines = beforeLines;
    }

}
