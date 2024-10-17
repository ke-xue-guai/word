package com.export.word.table;

import com.export.word.WordExportContext;
import org.docx4j.wml.Tbl;

import java.util.List;

public class PageTable extends TitleTable {

    @Override
    public void toElement(WordExportContext ctx, List<Object> content) {
        if (this.title != null) {
            this.title.toElement(ctx, content);
        }
        if (this.content != null) {
            if (content.get(content.size() - 1) instanceof Tbl) {
                this.content.setAbsolute(true);
            }
            this.content.toElement(ctx, content);
        }
    }

}
