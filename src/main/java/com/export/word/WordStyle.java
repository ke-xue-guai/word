package com.export.word;

import org.docx4j.wml.SectPr;

public class WordStyle extends Single<SectPr> {

    private WordSize size;
    private WordMargins margins;
    private Boolean isFixRowHigh = true;

    @Override
    public SectPr create() {
        SectPr pr = new SectPr();
        // 页面的边距
        if (this.margins != null) {
            pr.setPgMar(this.margins.get());
        }
        // 纸张的属性
        if (this.size != null) {
            pr.setPgSz(this.size.get());
        }
        return pr;
    }

    public WordSize getSize() {
        return size;
    }

    public void setSize(WordSize size) {
        this.size = size;
    }

    public WordMargins getMargins() {
        return margins;
    }

    public void setMargins(WordMargins margins) {
        this.margins = margins;
    }

    public Boolean getFixRowHigh() {
        return isFixRowHigh;
    }

    public void setFixRowHigh(Boolean fixRowHigh) {
        isFixRowHigh = fixRowHigh;
    }
}
