package com.export.word;

import org.docx4j.wml.HpsMeasure;

import java.math.BigInteger;

public class FontSize extends Single<HpsMeasure> {

    public static final FontSize DEFAULT = new FontSize(10L);

    private Long value;

    public FontSize() {

    }

    public FontSize(Long value) {
        this.value = value;
    }

    @Override
    public HpsMeasure create() {
        HpsMeasure size = new HpsMeasure();
        if (this.value != null) {
            size.setVal(BigInteger.valueOf(value));
        }
        return size;
    }

    public Long getValue() {
        return value;
    }

    public void setValue(Long value) {
        this.value = value;
    }

}
