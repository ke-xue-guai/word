package com.export.word;

import org.docx4j.wml.CTVerticalJc;
import org.docx4j.wml.STVerticalJc;

/**
 * 垂直对齐方式
 */
public class VerticalAlign extends Single<CTVerticalJc> {

    private Type type;

    public VerticalAlign() {

    }

    public VerticalAlign(Type type) {
        this.type = type;
    }

    @Override
    public CTVerticalJc create() {
        CTVerticalJc vertical = new CTVerticalJc();
        if (this.type != null) {
            vertical.setVal(STVerticalJc.valueOf(String.valueOf(this.type)));
        }
        return vertical;
    }


    public Type getType() {
        return type;
    }

    public void setType(Type type) {
        this.type = type;
    }

    public enum Type {
        TOP,
        CENTER,
        BOTH,
        BOTTOM;
    }

}
