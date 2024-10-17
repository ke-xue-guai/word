package com.export.word;

import org.docx4j.wml.STPageOrientation;
import org.docx4j.wml.SectPr;

import java.math.BigInteger;

/**
 * 文档的尺寸
 */
public class WordSize extends Single<SectPr.PgSz> {

    /**
     * 页面属性，0为纵向，1为横向
     */
    private short orient = 0;

    /**
     * 宽度 尺寸*28.35*20 单位转化 16443
     */
    private BigInteger pageWidth = BigInteger.valueOf((long) (29.7 * 28.35 * 20));

    /**
     * 高度 11907
     */
    private BigInteger pageHeight = BigInteger.valueOf((long) (21.0 * 28.35 * 20));

    @Override
    public SectPr.PgSz create() {
        SectPr.PgSz size = new SectPr.PgSz();
        if (orient == 1) {
            size.setOrient(STPageOrientation.LANDSCAPE);
        } else if (orient == 0) {
            size.setOrient(STPageOrientation.PORTRAIT);
        }
        size.setW(pageWidth);
        size.setH(pageHeight);
        return size;
    }

    public short getOrient() {
        return orient;
    }

    public void setOrient(short orient) {
        this.orient = orient;
    }

    public BigInteger getPageWidth() {
        return pageWidth;
    }

    public void setPageWidth(BigInteger pageWidth) {
        this.pageWidth = pageWidth;
    }

    public BigInteger getPageHeight() {
        return pageHeight;
    }

    public void setPageHeight(BigInteger pageHeight) {
        this.pageHeight = pageHeight;
    }
}
