package com.export.word;

import java.io.Serializable;
import java.util.List;

public abstract class Element implements Serializable {

    protected long height = 0;

    public abstract void toElement(WordExportContext ctx, List<Object> content);

    public long getHeight() {
        return this.height;
    }

    public void setHeight(long height) {
        this.height = height;
    }

}
