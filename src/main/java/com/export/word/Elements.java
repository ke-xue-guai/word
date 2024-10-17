package com.export.word;

import java.util.ArrayList;
import java.util.List;

public class Elements extends Element {

    private List<Element> list = new ArrayList<>();

    @Override
    public void toElement(WordExportContext ctx, List<Object> content) {
        for (Element element : list) {
            element.toElement(ctx, content);
        }
    }

    public void add(Element element) {
        this.list.add(element);
    }

}
