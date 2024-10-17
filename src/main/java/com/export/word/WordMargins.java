package com.export.word;

import org.docx4j.wml.SectPr;

import java.math.BigInteger;

/**
 * 文档的边距
 */
public class WordMargins extends Single<SectPr.PgMar> {

    public Long top = 220L;
    public Long left = 1440L;
    public Long right = 1440L;
    public Long bottom = 220L;
    public Long header = 0L;
    public Long footer = 0L;

    @Override
    public SectPr.PgMar create() {
        SectPr.PgMar mar = new SectPr.PgMar();
        mar.setTop(BigInteger.valueOf(this.top));
        mar.setLeft(BigInteger.valueOf(this.left));
        mar.setRight(BigInteger.valueOf(this.right));
        mar.setBottom(BigInteger.valueOf(this.bottom));
        mar.setHeader(BigInteger.valueOf(this.header));
        mar.setFooter(BigInteger.valueOf(this.footer));
        return mar;
    }

}
