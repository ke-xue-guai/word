package com.export.word.table;

import com.export.word.Element;
import com.export.word.WordExportContext;
import org.docx4j.wml.Tr;

import java.util.ArrayList;
import java.util.List;

public class TableRow extends Element {

    private List<TableCol> cols = new ArrayList<>();

    private TableRowStyle tableRowStyle;

    public void addCols(List<TableCol> tableCols) {
        this.cols.addAll(tableCols);
    }

    @Override
    public void toElement(WordExportContext ctx, List<Object> content) {
        // 添加行
        Tr row = new Tr();
        content.add(row);
        // 添加行样式
        if (this.tableRowStyle == null) {
            row.setTrPr(TableRowStyle.DEFAULT.get());
        } else {
            row.setTrPr(this.tableRowStyle.get());
        }
        // 添加单元格
        for (TableCol col : cols) {
            col.toElement(ctx, row.getContent());
        }
    }

    public void addCol(TableCol tableCol) {
        cols.add(tableCol);
    }

    public void addCol(Integer index, TableCol tableCol) {
        cols.add(index, tableCol);
    }

    public List<TableCol> getCols() {
        return cols;
    }

    public void setCols(List<TableCol> cols) {
        this.cols = cols;
    }

    @Override
    public long getHeight() {
        if (this.tableRowStyle == null) {
            return TableRowStyle.DEFAULT.getHeight();
        } else {
            return this.tableRowStyle.getHeight();
        }
    }

    public TableRowStyle getTableRowStyle() {
        return tableRowStyle;
    }

    public void setTableRowStyle(TableRowStyle tableRowStyle) {
        this.tableRowStyle = tableRowStyle;
    }

}
