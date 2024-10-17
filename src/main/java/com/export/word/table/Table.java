package com.export.word.table;

import com.export.word.Element;
import com.export.word.WordExportContext;
import com.export.word.style.WidthStyle;
import org.docx4j.wml.*;

import java.math.BigInteger;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Table extends Element {

    private List<TableRow> rows = new ArrayList<>();

    private List<WidthStyle> widthStyles = new ArrayList<>();

    /**
     * 是否是绝对定位
     */
    private Boolean absolute = false;

    /**
     * 自适应宽度，决定是否横向占满文档
     */
    private Boolean autoFit = true;


    public void addRow(TableRow row) {
        this.rows.add(row);
    }

    public void addAllRow(List<TableRow> rows) {
        this.rows.addAll(rows);
    }

    public void setWidthStyles(Integer... widthStyles) {
        for (int i = 0; i < widthStyles.length; i++) {
            WidthStyle widthStyle = new WidthStyle();
            widthStyle.setWidth(widthStyles[i]);
            widthStyle.setIndex(i);
            this.widthStyles.add(widthStyle);
        }
    }

    @Override
    public void toElement(WordExportContext ctx, List<Object> content) {
        if (this.rows.size() == 0) {
            return;
        }
        content.add(createTbl(ctx));
    }

    public List<TableRow> getRows() {
        return rows;
    }

    public void setRows(List<TableRow> rows) {
        this.rows = rows;
    }

    public List<WidthStyle> getWidthStyles() {
        return widthStyles;
    }

    public void setWidthStyles(List<WidthStyle> widthStyles) {
        this.widthStyles = widthStyles;
    }

    public boolean getAbsolute() {
        return this.absolute;
    }

    public void setAbsolute(boolean absolute) {
        this.absolute = absolute;
    }

    private Tbl createTbl(WordExportContext ctx) {
        // 填充单元格
        this.fillContinueCol();

        Tbl tbl = new Tbl();
        // 填充单元格列宽
        TblGrid tblGrid = this.getTblGrid(ctx);
        tbl.setTblGrid(tblGrid);
        // 添加行数据
        for (TableRow row : this.getRows()) {
            row.toElement(ctx, tbl.getContent());
        }
        // 添加表格样式
        TblPr tblPr = createTblPr(ctx);
        tbl.setTblPr(tblPr);
        return tbl;
    }

    private TblPr createTblPr(WordExportContext ctx) {
        // 添加样式
        TblPr tblPr = new TblPr();
        // 是否是绝对定位
        if (this.absolute) {
            CTTblPPr ctTblPPr = new CTTblPPr();
            ctTblPPr.setTblpX(BigInteger.valueOf(ctx.wordMargins.left));
            ctTblPPr.setTblpY(BigInteger.TEN);
            ctTblPPr.setHorzAnchor(STHAnchor.PAGE);
            ctTblPPr.setLeftFromText(BigInteger.valueOf(180));
            ctTblPPr.setRightFromText(BigInteger.valueOf(180));
            ctTblPPr.setVertAnchor(STVAnchor.TEXT);
            tblPr.setTblpPr(ctTblPPr);
            // 环绕类型
            CTTblOverlap ctTblOverlap = new CTTblOverlap();
            ctTblOverlap.setVal(STTblOverlap.NEVER);
            tblPr.setTblOverlap(ctTblOverlap);
        }
        // 设置边距
        setTblMar(tblPr);
        return tblPr;
    }

    /**
     * 边距
     */
    private void setTblMar(TblPr tblPr) {
        TblWidth zero = new TblWidth();
        zero.setW(BigInteger.ZERO);
        zero.setType(TblWidth.TYPE_DXA);
        // 设置单元格边距
        CTTblCellMar margin = new CTTblCellMar();
        margin.setTop(zero);
        margin.setBottom(zero);
        margin.setLeft(zero);
        margin.setRight(zero);
        tblPr.setTblCellMar(margin);
        // 表格宽度
        tblPr.setTblW(zero);
        // 表格缩进
        tblPr.setTblInd(zero);
        // 布局类型
        CTTblLayoutType ctTblLayoutType = new CTTblLayoutType();
        ctTblLayoutType.setType(STTblLayoutType.FIXED);
        tblPr.setTblLayout(ctTblLayoutType);
    }

    /**
     * 设置列宽度
     */
    private TblGrid getTblGrid(WordExportContext ctx) {
        int maxColNumber = this.getMaxColNumber();
        // 获取页面宽度
        long pageWidth = ctx.pageWidth.longValue();
        // 计算表格宽度
        long tableWidth = pageWidth - ctx.wordMargins.left - ctx.wordMargins.right;
        // 表格布局
        TblGrid tblGrid = new TblGrid();
        List<Long> colWidthList = new ArrayList<>();
        // 是否是自适应，自适应宽度，应先计算单元格所占的比例，然后求出实际的宽度
        if (this.autoFit) {
            // 补充单元格的宽度设置
            while (maxColNumber > this.widthStyles.size()) {
                WidthStyle width = new WidthStyle();
                width.setWidth(100);
                this.widthStyles.add(width);
            }
            // 计算总宽度
            long total = 0;
            for (WidthStyle widthStyle : this.widthStyles) {
                total += widthStyle.getWidth();
            }
            // 计算总共使用宽度
            long used = 0;
            for (WidthStyle widthStyle : this.widthStyles) {
                long computed = tableWidth * widthStyle.getWidth() / total / 72 * 72;
                used += computed;
                colWidthList.add(computed);
            }
            // 计算缺失宽度
            long lose = tableWidth - used;
            if (lose > 0) {
                colWidthList.set(0, (colWidthList.get(0) + lose) / 72 * 72);
            }
        }
        // 否则直接设置单元格的宽度
        else {
            // 补充单元格的宽度设置
            while (maxColNumber > this.widthStyles.size()) {
                WidthStyle width = new WidthStyle();
                width.setWidth(1000);
                this.widthStyles.add(width);
            }
            for (WidthStyle widthStyle : this.widthStyles) {
                colWidthList.add(Long.valueOf(widthStyle.getWidth()));
            }
        }

        Map<Integer, Long> colWidthMap = new HashMap<>();

        // 单元格宽度
        List<TblGridCol> gridCol = tblGrid.getGridCol();
        for (int i = 0; i < colWidthList.size(); i++) {
            Long width = colWidthList.get(i);
            TblGridCol tblGridCol = new TblGridCol();
            tblGridCol.setW(BigInteger.valueOf(width));
            gridCol.add(tblGridCol);
            colWidthMap.put(i, width);
        }

        // 设置单元格宽度，需要先设置gridCol在设置fix
        for (TableRow row : this.rows) {
            List<TableCol> cols = row.getCols();
            for (int i = 0; i < cols.size(); i++) {
                TableCol col = cols.get(i);
                long width = colWidthMap.get(i);
                for (int offset = 1; offset < col.getColspan(); offset++) {
                    width += colWidthMap.get(i + offset);
                }
                col.setWidth(width);
            }
        }
        return tblGrid;
    }

    /**
     * 填充空白单元格
     */
    public void fillContinueCol() {
        // 删除全空行
        List<Integer> removeRows = new ArrayList<>();
        for (int x = 0; x < this.rows.size(); x++) {
            if (this.rows.get(x).getCols().isEmpty()) {
                removeRows.add(0, x);
            }
        }
        for (int removeRow : removeRows) {
            this.rows.remove(removeRow);
        }
        // 是否存在完全合并的行，rowspan相同并且大于1
        Map<Integer, Integer> minRowspanMap = new HashMap<>();
        for (int x = 0; x < this.rows.size(); x++) {
            TableRow row = this.rows.get(x);
            int min = 999;
            int max = 0;
            for (TableCol col : row.getCols()) {
                int rowspan = col.getRowspan();
                if (min > rowspan) {
                    min = rowspan;
                }
                if (max < rowspan) {
                    max = rowspan;
                }
            }
            // 最大值和最小值相同
            if (min == max && min > 1) {
                for (TableCol col : row.getCols()) {
                    col.setRowspan(1);
                }
                int index_X = x;
                int totalRowSpan = 0;
                while (index_X - 1 >= 0) {
                    TableRow preRow = rows.get(index_X - 1);
                    for (TableCol col : preRow.getCols()) {
                        // 上级跨行包含本行
                        if (col.getRowspan() > min + totalRowSpan) {
                            col.setRowspan(col.getRowspan() - min + 1);
                            if (col.getRowspan() == 1) {
                                // 跨行为1时更新跨行合并
                                col.setVMerge(null);
                            }
                        }
                    }
                    totalRowSpan += min - 1;
                    index_X--;
                }
                if (row.getTableRowStyle() == null) {
                    row.setTableRowStyle(new TableRowStyle());
                    row.getTableRowStyle().setFixHeight(true);
                }
                row.getTableRowStyle().setHeight(row.getHeight() * min);
            }
            minRowspanMap.put(x, min);
        }
        Integer maxColNumber = getMaxColNumber();
        // 补充被跨行合并的单元格
        for (int x = 0; x < this.rows.size(); x++) {
            TableRow row = this.rows.get(x);
            List<TableCol> cols = row.getCols();
            for (int y = 0; y < cols.size(); y++) {
                TableCol col = cols.get(y);
                int rowspan = col.getRowspan();
                if (rowspan > 1) {
                    // 设置合并行单元格属性
                    col.setVMerge(TableCol.VMerge.RESTART);
                    // 下一行总列数
                    int nextRowTotalColSpan = 0;
                    TableRow nextRow = this.rows.get(x + 1);
                    for (int i = 0; i < nextRow.getCols().size(); i++) {
                        nextRowTotalColSpan += nextRow.getCols().get(i).getColspan();
                    }
                    // 下一行总列数等于最大列数，说明已填充
                    if (nextRowTotalColSpan == maxColNumber) {
                        continue;
                    }
                    // 获取当前单元格之前跨列合并值
                    int totalColSpan = 0;
                    for (int i = 0; i <= y; i++) {
                        totalColSpan += cols.get(i).getColspan() - 1;
                    }
                    for (int i = 1; i < rowspan; ) {
                        TableCol continueCol = new TableCol();
                        continueCol.setPositionY(col.getPositionY());
                        continueCol.setVMerge(TableCol.VMerge.CONTINUE);
                        continueCol.setColspan(col.getColspan());
                        continueCol.setTableColStyle(col.getTableColStyle());
                        // 添加到下一行
                        TableRow next = this.rows.get(x + i);
                        int positionY = y + totalColSpan;
                        boolean isAdd = false;
                        for (int j = 0; j < next.getCols().size(); j++) {
                            if (next.getCols().get(j).getPositionY() > positionY) {
                                next.getCols().add(j, continueCol);
                                isAdd = true;
                                break;
                            }
                        }
                        if (!isAdd) {
                            next.getCols().add(next.getCols().size(), continueCol);
                        }
                        i += minRowspanMap.get(x + i);
                    }
                }
            }
        }

        // 列上的所有单元格都有大于1的colspan属性，应忽略其小于最小值的部分
        Map<Integer, List<TableCol>> colIndexMap = new HashMap<>();
        for (TableRow row : this.rows) {
            int index = 0;
            for (TableCol col : row.getCols()) {
                int colspan = col.getColspan();
                if (colspan > 1) {
                    List<TableCol> cols = colIndexMap.computeIfAbsent(index, k -> new ArrayList<>());
                    cols.add(col);
                }
                index += colspan;
            }
        }
        for (Map.Entry<Integer, List<TableCol>> entry : colIndexMap.entrySet()) {
            List<TableCol> cols = entry.getValue();
            if (cols.size() == this.rows.size()) {
                int min = 999;
                for (TableCol col : cols) {
                    int colspan = col.getColspan();
                    if (min > colspan) {
                        min = colspan;
                    }
                }
                min -= 1;
                for (TableCol col : cols) {
                    int colspan = col.getColspan();
                    col.setColspan(colspan - min);
                }
            }
        }
    }

    /**
     * 获取最大列数
     */
    public Integer getMaxColNumber() {
        int max = 0;
        //获取最大列的下标
        for (TableRow row : this.rows) {
            List<TableCol> cols = row.getCols();
            for (int i = 0; i < cols.size(); i++) {
                TableCol col = cols.get(i);
                int number = i + col.getColspan();
                if (max < number) {
                    max = number;
                }
            }
        }
        return max;
    }

    @Override
    public long getHeight() {
        int rst = 0;
        for (TableRow row : this.rows) {
            rst += row.getHeight();
        }
        return rst;
    }

    public Boolean getAutoFit() {
        return this.autoFit;
    }

    public void setAutoFit(Boolean autoFit) {
        this.autoFit = autoFit;
    }

}
