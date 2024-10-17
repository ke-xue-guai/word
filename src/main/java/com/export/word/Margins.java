package com.export.word;

import org.docx4j.jaxb.Context;
import org.docx4j.wml.TblWidth;
import org.docx4j.wml.TcMar;

import java.math.BigInteger;


public class Margins extends Single<TcMar> {

    public Long top;

    public Long left;

    public Long right;

    public Long bottom;

    @Override
    public TcMar create() {
        TcMar tcMar = new TcMar();

        if (this.top != null) {
            tcMar.setTop(width(this.top));
        }
        if (this.left != null) {
            tcMar.setLeft(width(this.left));
        }
        if (this.right != null) {
            tcMar.setRight(width(this.right));
        }
        if (this.bottom != null) {
            tcMar.setBottom(width(this.bottom));
        }

        return tcMar;
    }

    private TblWidth width(long v) {
        TblWidth width = Context.getWmlObjectFactory().createTblWidth();
        width.setW(BigInteger.valueOf(v));
        width.setType(TblWidth.TYPE_DXA);
        return width;
    }

}
