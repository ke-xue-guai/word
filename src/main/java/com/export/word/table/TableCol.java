package com.export.word.table;

import com.export.word.*;
import org.docx4j.wml.TblWidth;
import org.docx4j.wml.Tc;
import org.docx4j.wml.TcPr;
import org.docx4j.wml.TcPrInner;

import java.math.BigInteger;
import java.util.List;

public class TableCol extends Element {

    /**
     * 行
     */
    private int positionX;

    /**
     * 列
     */
    private int positionY;

    /**
     * 单元格所对应的值
     */
    private String value;

    /**
     * 合并列
     */
    private int colspan = 1;

    /**
     * 合并行
     */
    private int rowspan = 1;

    /**
     * 私有属性
     */
    private VMerge vMerge;

    /**
     * 私有属性
     */
    private long width;

    /**
     * 单元格样式
     */
    private TableColStyle tableColStyle;

    /**
     * 字体
     */
    private ParagraphTextStyle paragraphTextStyle;

    @Override
    public void toElement(WordExportContext ctx, List<Object> content) {
        // 添加单元格值
        Tc col = new Tc();
        content.add(col);
        // 添加单元格属性
        TextAlign textAlign = null;
        if (tableColStyle != null) {
            ParagraphStyle paragraphStyle = tableColStyle.getParagraphStyle();
            if (paragraphStyle != null) {
                textAlign = paragraphStyle.getTextAlign();
            }
        }
        if (paragraphTextStyle == null) {
            paragraphTextStyle = ParagraphTextStyle.DEFAULT;
        }
        Paragraph paragraph = new Paragraph(this.value, paragraphTextStyle.getFont(), paragraphTextStyle.getFontSize(), textAlign);
        if (this.tableColStyle != null) {
            col.setTcPr(this.tableColStyle.create());
            paragraph.setParagraphStyle(this.tableColStyle.getParagraphStyle());
        } else {
            col.setTcPr(TableColStyle.DEFAULT.create());
            paragraph.setParagraphStyle(TableColStyle.DEFAULT.getParagraphStyle());
        }
        TcPr tcPr = col.getTcPr();
        if (this.vMerge != null) {
            TcPrInner.VMerge vMerge = new TcPrInner.VMerge();
            vMerge.setVal(this.vMerge.getValue());
            tcPr.setVMerge(vMerge);
        }
        if (this.colspan > 1) {
            TcPrInner.GridSpan gridSpan = new TcPrInner.GridSpan();
            gridSpan.setVal(BigInteger.valueOf(this.colspan));
            tcPr.setGridSpan(gridSpan);
        }
        TblWidth width = new TblWidth();
        width.setType(TblWidth.TYPE_DXA);
        width.setW(BigInteger.valueOf(this.width));
        tcPr.setTcW(width);
        // todo 行间距设置1.5倍行距
        paragraph.toElement(ctx, col.getContent());
    }


    public int getPositionX() {
        return positionX;
    }

    public void setPositionX(int positionX) {
        this.positionX = positionX;
    }

    public int getPositionY() {
        return positionY;
    }

    public void setPositionY(int positionY) {
        this.positionY = positionY;
    }

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }

    public int getColspan() {
        return colspan;
    }

    public void setColspan(int colspan) {
        this.colspan = colspan;
    }

    public int getRowspan() {
        return rowspan;
    }

    public void setRowspan(int rowspan) {
        this.rowspan = rowspan;
    }

    public TableColStyle getTableColStyle() {
        return tableColStyle;
    }

    public void setTableColStyle(TableColStyle tableColStyle) {
        this.tableColStyle = tableColStyle;
    }

    public VMerge getVMerge() {
        return vMerge;
    }

    public void setVMerge(VMerge vMerge) {
        this.vMerge = vMerge;
    }

    public long getWidth() {
        return width;
    }

    public void setWidth(long width) {
        this.width = width;
    }

    public ParagraphTextStyle getParagraphTextStyle() {
        return paragraphTextStyle;
    }

    public void setParagraphTextStyle(ParagraphTextStyle paragraphTextStyle) {
        this.paragraphTextStyle = paragraphTextStyle;
    }

    public enum VMerge {

        RESTART("restart"),
        CONTINUE("continue");

        private String value;

        VMerge(String value) {
            this.value = value;
        }

        public String getValue() {
            return value;
        }

    }


}
