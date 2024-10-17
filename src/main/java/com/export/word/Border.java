package com.export.word;

import org.docx4j.wml.CTBorder;
import org.docx4j.wml.STBorder;

import java.math.BigInteger;

public class Border extends Single<CTBorder> {

    private static final String DEFAULT_COLOR = "000000";
    private static final BigInteger DEFAULT_SIZE = BigInteger.valueOf(4);
    private static final BigInteger DEFAULT_SPACE = BigInteger.valueOf(0);

    private Type type = Type.NIL;
    private String color = DEFAULT_COLOR;
    private BigInteger size = DEFAULT_SIZE;
    private BigInteger space = DEFAULT_SPACE;

    public void setType(Type type) {
        this.type = type;
    }

    public void setColor(String color) {
        this.color = color;
    }

    public void setSize(BigInteger size) {
        this.size = size;
    }

    public void setSpace(BigInteger space) {
        this.space = space;
    }

    public Type getType() {
        return type;
    }

    public String getColor() {
        return color;
    }

    public BigInteger getSize() {
        return size;
    }

    public BigInteger getSpace() {
        return space;
    }

    @Override
    public CTBorder create() {
        CTBorder instance = new CTBorder();
        if (this.type != null) {
            instance.setColor(this.color);
            instance.setSz(this.size);
            instance.setSpace(this.space);
            instance.setVal(STBorder.valueOf(String.valueOf(this.type)));
        }
        return instance;
    }

    /**
     * 边框样式：无、有边框、加粗边框
     */
    public enum Type {
        NIL, NONE, THICK, SINGLE, DOUBLE
    }

}
