package com.export.word.table;

import com.export.word.*;
import org.docx4j.wml.TcPr;

public class TableColStyle extends Single<TcPr> {

    private Margins margins;
    private TableColBorders tableColBorders;
    private VerticalAlign verticalAlign;

    /**
     * 单元格设置段落属性
     */
    private ParagraphStyle paragraphStyle;

    public static final TableColStyle DEFAULT = new TableColStyle();

    static {
        TextAlign defaultTextAlign = new TextAlign();
        defaultTextAlign.setType(TextAlign.Type.LEFT);
        VerticalAlign defaultVerticalAlign = new VerticalAlign();
        defaultVerticalAlign.setType(VerticalAlign.Type.CENTER);
        DEFAULT.setVerticalAlign(defaultVerticalAlign);

        ParagraphStyle defaultParagraphStyle = new ParagraphStyle();
        defaultParagraphStyle.setTextAlign(defaultTextAlign);
        defaultParagraphStyle.setSpacing(new Spacing(0L));
        DEFAULT.setParagraphStyle(defaultParagraphStyle);
    }

    @Override
    public TcPr create() {
        TcPr tcPr = new TcPr();
        // 边距
        if (this.margins != null) {
            tcPr.setTcMar(margins.get());
        }
        // 边框
        if (this.tableColBorders != null) {
            tcPr.setTcBorders(this.tableColBorders.get());
        }
        // 垂直对齐方式
        if (this.verticalAlign != null) {
            tcPr.setVAlign(this.verticalAlign.get());
        }
        return tcPr;
    }

    public Margins getMargins() {
        return margins;
    }

    public void setMargins(Margins margins) {
        this.margins = margins;
    }

    public TableColBorders getTableColBorders() {
        return tableColBorders;
    }

    public void setTableColBorders(TableColBorders tableColBorders) {
        this.tableColBorders = tableColBorders;
    }

    public VerticalAlign getVerticalAlign() {
        return verticalAlign;
    }

    public void setVerticalAlign(VerticalAlign verticalAlign) {
        this.verticalAlign = verticalAlign;
    }

    public ParagraphStyle getParagraphStyle() {
        return paragraphStyle;
    }

    public void setParagraphStyle(ParagraphStyle paragraphStyle) {
        this.paragraphStyle = paragraphStyle;
    }

}
