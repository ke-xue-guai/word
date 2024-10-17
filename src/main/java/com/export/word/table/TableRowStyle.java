package com.export.word.table;

import com.export.word.Single;
import org.docx4j.model.properties.table.tr.TrHeight;
import org.docx4j.wml.*;

import java.math.BigInteger;

public class TableRowStyle extends Single<TrPr> {

    private long height = 280;
    private boolean isFixHeight = true;

    public static final TableRowStyle DEFAULT = new TableRowStyle();

    @Override
    public TrPr create() {
        TrPr trPr = new TrPr();
        CTHeight height = new CTHeight();
        height.setVal(BigInteger.valueOf(this.height));
        // 设置固定行高
        if (this.isFixHeight) {
            height.setHRule(STHeightRule.EXACT);
        }
        TrHeight trHeight = new TrHeight(height);
        trHeight.set(trPr);
        return trPr;
    }

    public long getHeight() {
        return height;
    }

    public void setHeight(long height) {
        this.height = height;
    }

    public boolean isFixHeight() {
        return isFixHeight;
    }

    public void setFixHeight(boolean fixHeight) {
        isFixHeight = fixHeight;
    }

}
