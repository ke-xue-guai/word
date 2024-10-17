package com.export.word.table;

import com.export.word.Element;

/**
 * 带标题的表
 */
public abstract class TitleTable extends Element {

    /**
     * 标题
     */
    protected Element title;

    /**
     * 内容
     */
    protected Table content;

    public Element getTitle() {
        return title;
    }

    public void setTitle(Element title) {
        this.title = title;
    }

    public Table getContent() {
        return content;
    }

    public void setContent(Table content) {
        this.content = content;
    }

}
