package com.export.word;

import org.docx4j.wml.Jc;
import org.docx4j.wml.JcEnumeration;

/**
 * 水平对齐方式
 */
public class TextAlign extends Single<Jc> {

    private Type type;

    public TextAlign() {

    }

    public TextAlign(Type type) {
        this.type = type;
    }

    @Override
    public Jc create() {
        Jc vertical = new Jc();
        if (this.type != null) {
            vertical.setVal(JcEnumeration.valueOf(String.valueOf(this.type)));
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
        LEFT,
        CENTER,
        RIGHT
    }

}
