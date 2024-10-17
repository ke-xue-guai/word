package com.export.word.table;

import com.export.word.WordExportContext;
import com.export.word.WordMargins;
import com.export.word.style.WidthStyle;
import com.export.word.utils.StringUtils;

import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;

/**
 * 大表
 * <p>
 * 大表需要先拆成多个页面表
 */
public class BigTable extends TitleTable {


    /**
     * 重复列数
     */
    private Integer repeatColumn = 0;

    /**
     * 列分页属性
     */
    private String splitTableProperty;

    /**
     * 重复行
     */
    private Integer repeatRow = 0;

    /**
     * 分页类型 0为Z 1为N,默认为0
     */
    private Integer pageType = 0;

    /**
     * 是否自适应行高
     */
    private Boolean isAutoFit = false;


    @Override
    public void toElement(WordExportContext ctx, List<Object> content) {
        List<PageTable> tables = splitTable(ctx);
        for (PageTable table : tables) {
            table.toElement(ctx, content);
        }
    }

    /**
     * 拆表
     * 按行拆/按列拆
     * N型 Z型
     * 固定列 4  按列拆表处理固定列{4~6},{8~12}{13~14}
     */
    private List<PageTable> splitTable(WordExportContext ctx) {
        List<PageTable> rst = new ArrayList<>();
        // 填充初始列宽度
        if (this.content.getWidthStyles() == null || this.content.getWidthStyles().size() == 0) {
            List<WidthStyle> widthStyles = new ArrayList<>();
            for (Integer index = 0; index < this.content.getMaxColNumber(); index++) {
                WidthStyle widthStyle = new WidthStyle();
                widthStyle.setIndex(index);
                widthStyle.setWidth(10);
                widthStyles.add(widthStyle);
            }
            this.content.setWidthStyles(widthStyles);
        }
        // 填充单元格 垂直
        this.content.fillContinueCol();
        // 默认先跨列拆
        if (this.pageType == 0) {
            // 按列拆
            List<Table> tablesByCol = this.splitTableByCol(this.content, splitTableProperty);
            // 按行拆
            for (Table tablePage : tablesByCol) {
                List<Table> tablesByRow = this.splitTableByRow(tablePage, ctx);
                rst.addAll(convert(tablesByRow, ctx));
            }
        } else {
            List<Table> tablesByRow = this.splitTableByRow(this.content, ctx);
            for (Table tablePage : tablesByRow) {
                List<Table> tablesByCol = this.splitTableByCol(tablePage, splitTableProperty);
                rst.addAll(convert(tablesByCol, ctx));
            }
        }
        clear(rst);
        return rst;
    }

    /**
     * 清除被合并单元格
     */
    private void clear(List<PageTable> tables) {
        for (PageTable pageTable : tables) {
            Table table = pageTable.getContent();
            for (TableRow row : table.getRows()) {
                List<TableCol> cols = row.getCols();
                // 进行判断和操作
                cols.removeIf(next -> next.getVMerge() == TableCol.VMerge.CONTINUE);
            }
        }
    }

    /**
     * 拆表后Table转大表，判断是否开启自适应
     */
    private List<PageTable> convert(List<Table> tables, WordExportContext ctx) {
        List<PageTable> rst = new ArrayList<>();
        for (Table table : tables) {
//            this.setAutoFit(this, table, ctx);
            PageTable pageTable = new PageTable();
            pageTable.setTitle(this.title);
            pageTable.setContent(table);
            rst.add(pageTable);
        }
        return rst;
    }

    /**
     * 按行拆表(新)
     */
    private List<Table> splitTableByRow(Table sourceTable, WordExportContext ctx) {
        List<Table> rst = new ArrayList<>();
        List<TableRow> allRows = sourceTable.getRows();
        int allRowSize = allRows.size();
        // 获取重复行
        List<TableRow> repeatRows = getRepeatRows(sourceTable);
        int repeatNum = repeatRows.size();
        // 整个表高度
        long heights = sourceTable.getHeight();
        long effectiveHeight = getEffectiveHeight(this, ctx);
        // 获取重复行高度
        int repeatH = 0;
        for (TableRow repeatRow : repeatRows) {
            repeatH += repeatRow.getHeight();
        }
        if (repeatNum == allRowSize || heights <= effectiveHeight) {
            // 固定表 || 有效高度大于表高度，不需要拆表
            rst.add(sourceTable);
            return rst;
        }
        // 有效高度-重复行高度等于剩余高度
        long residueH = effectiveHeight - repeatH;
        if (residueH < 0) {
            throw new RuntimeException("剩余高度不足");
        }
        // 计算需要拆成几张表的行索引
        List<Integer> indexPageByRow = new ArrayList<>();
        int usedNum = repeatNum;
        while (usedNum < sourceTable.getRows().size()) {
            int sum = 0;
            int index = usedNum;
            for (; index < sourceTable.getRows().size(); index++) {
                sum += sourceTable.getRows().get(index).getHeight();
                if (sum >= residueH) {
                    //记录 行index
                    indexPageByRow.add(index - 1);
                    usedNum = index - 1;
                    break;
                }
            }
            if (index == sourceTable.getRows().size()) {
                indexPageByRow.add(index);
                break;
            }
        }
        // 重新设置被使用的行数量
        usedNum = repeatNum;
        for (Integer indexRow : indexPageByRow) {
            Table newTable = new Table();
            // 塞重复行
            newTable.addAllRow(repeatRows);
            // 处理数据
            List<TableRow> tableRows = sourceTable.getRows().subList(usedNum, indexRow);
            newTable.addAllRow(tableRows);
            newTable.setWidthStyles(sourceTable.getWidthStyles());
            rst.add(newTable);
            usedNum = indexRow;
        }
        return rst;
    }

    /**
     * 获取重复行(新)
     */
    private List<TableRow> getRepeatRows(Table sourceTable) {
        return sourceTable.getRows().subList(0, repeatRow);
    }

    private void addCol(Table table, int x, TableCol col) {
        while (table.getRows().size() <= x) {
            table.getRows().add(new TableRow());
        }
        TableRow row = table.getRows().get(x);
        if (row == null) {
            row = new TableRow();
            table.addRow(row);
        }
        row.addCol(col);
    }

    /**
     * 按列拆(新)
     */
    private List<Table> splitTableByCol(Table sourceTable, String splitTableProperty) {
        List<Table> rst = new ArrayList<>();
        if (StringUtils.isBlank(splitTableProperty)) {
            rst.add(sourceTable);
            return rst;
        }
        // 拆成多张表 {4~6}{8~12,13~14}
        splitTableProperty = splitTableProperty.substring(1, splitTableProperty.length() - 1);
        String[] tableProperties = splitTableProperty.split("}\\{");
        // 初始化表索引
        List<List<Integer>> tableIndexes = new ArrayList<>();
        for (String tableProperty : tableProperties) {
            String[] ranges = tableProperty.split(",");
            List<Integer> indexes = new ArrayList<>();
            for (String range : ranges) {
                String[] splits = range.split("~");
                int start = Integer.parseInt(splits[0]);
                int end = Integer.parseInt(splits[1]);
                for (int i = start; i <= end; i++) {
                    indexes.add(i);
                }
            }
            tableIndexes.add(indexes);
            rst.add(new Table());
        }
        List<TableRow> rows = sourceTable.getRows();
        for (int x = 0; x < rows.size(); x++) {
            TableRow row = rows.get(x);
            List<TableCol> cols = row.getCols();
            if (cols.stream().allMatch(c -> c.getVMerge() == TableCol.VMerge.CONTINUE)) {
                continue;
            }
            int y = 0;
            for (int index = 0; index < cols.size(); index++) {
                TableCol col = cols.get(index);
                int step = col.getColspan();
                // 如果是被合并的列，跳过
                if (col.getVMerge() == TableCol.VMerge.CONTINUE) {
                    y += step;
                    continue;
                }
                // 添加重复列
                if (y < repeatColumn) {
                    for (int i = 0; i < tableIndexes.size(); i++) {
                        addCol(rst.get(i), x, col);
                    }
                    y += step;
                    continue;
                }
                // 拆表
                while (true) {
                    boolean hasAppend = false;
                    for (int i = 0; i < tableIndexes.size(); i++) {
                        List<Integer> indexes = tableIndexes.get(i);
                        if (indexes.contains(y)) {
                            int colspan = col.getColspan();
                            if (colspan == 1) {
                                addCol(rst.get(i), x, col);
                                hasAppend = true;
                                break;
                            } else {
                                int colSplitLen = 0;
                                for (int p = 0; p < colspan; p++) {
                                    if (indexes.contains(y + p)) {
                                        colSplitLen++;
                                    }
                                }
                                if (colSplitLen == colspan) {
                                    addCol(rst.get(i), x, col);
                                    hasAppend = true;
                                    break;
                                } else {
                                    // 拆单元格
                                    TableCol clone = new TableCol();
                                    clone.setValue(col.getValue());
                                    clone.setColspan(colSplitLen);
                                    clone.setRowspan(col.getRowspan());
                                    clone.setTableColStyle(col.getTableColStyle());
                                    clone.setPositionY(y);
                                    addCol(rst.get(i), x, clone);
                                    boolean isAdd = false;
                                    for (int j = 1; j < cols.size(); j++) {
                                        if (cols.get(j).getPositionY() > y) {
                                            cols.add(j - 1, clone);
                                            isAdd = true;
                                            break;
                                        }
                                    }
                                    if (!isAdd) {
                                        cols.add(cols.size() - 1, clone);
                                    }
                                    // 更新单元格跨列属性
//                                    y += colSplitLen;
                                    step = colSplitLen;
                                    col.setColspan(colspan - colSplitLen);
                                    // 设置单元格宽度
                                    col.setPositionY(step + col.getPositionY());
                                }
                            }
                        } else {
                            hasAppend = true;
                        }
                    }
                    if (hasAppend) {
                        break;
                    }
                    if (col.getColspan() > 1) {
                        // 处理跨列值为5， 拆分为0，跨列为3   3，跨列为2的情况，
                        // 拆表为0~3（0，1，2）， 4~...（4） 此时跨列为1
                        y++;
                        col.setColspan(col.getColspan() - 1);
                    } else {
                        break;
                    }
                }
                y += step;
            }
        }
        // 设置行高
        for (int i = 0; i < sourceTable.getRows().size(); i++) {
            for (Table table : rst) {
                for (int j = 0; j < sourceTable.getRows().size(); j++) {
                    table.getRows().get(j).setTableRowStyle(sourceTable.getRows().get(j).getTableRowStyle());
                }
            }
        }
        // 处理下标，由重复列引起的下标偏移
        for (int i = 1; i < tableIndexes.size(); i++) {
            List<Integer> integers = tableIndexes.get(i);
            List<Integer> preIndex = tableIndexes.get(i - 1);
            for (TableRow row : sourceTable.getRows()) {
                for (TableCol col : row.getCols()) {
                    if (integers.contains(col.getPositionY())) {
                        col.setPositionY(col.getPositionY() - preIndex.get(preIndex.size() - 1)  + repeatColumn);
                    }
                }
            }
        }
        // 处理下标，设置列宽度
        for (int i = 0; i < tableIndexes.size(); i++) {
            List<Integer> integers = tableIndexes.get(i);
            List<WidthStyle> widthStyles = new ArrayList<>(sourceTable.getWidthStyles().subList(0, repeatColumn));
            for (Integer integer : integers) {
                widthStyles.add(sourceTable.getWidthStyles().get(integer));
            }
            rst.get(i).setWidthStyles(widthStyles);
        }
        return rst;
    }

    public void setSplitTableProperty(String splitTableProperty) {
        this.splitTableProperty = splitTableProperty;
    }

    public void setRepeatColumn(Integer repeatColumn) {
        this.repeatColumn = repeatColumn;
    }

    public Boolean getAutoFit() {
        return isAutoFit;
    }

    public void setAutoFit(Boolean autoFit) {
        isAutoFit = autoFit;
    }

    public void setRepeatRow(Integer repeatRow) {
        this.repeatRow = repeatRow;
    }

    public Integer getPageType() {
        return pageType;
    }

    public void setPageType(Integer pageType) {
        this.pageType = pageType;
    }

    /**
     * 获取有效行高
     */
    private long getEffectiveHeight(BigTable bigTable, WordExportContext ctx) {
        WordMargins wordMargins = ctx.wordMargins;
        BigInteger pageH = ctx.pageHeight;
        long topMarH = wordMargins.top;
        long bottomMarH = wordMargins.bottom;
        long headerMargin = wordMargins.header;
        long footerMargin = wordMargins.footer;
        int tableRows = bigTable.getContent().getRows().size();
        long titleHeight = 0;
        if (bigTable.title != null) {
            titleHeight = bigTable.title.getHeight();
        }
        // 每行上下边框大小为8 = 4*2
        int borderHigh = tableRows * 8;
        // 整个页面的去除边界的有效高度
        return pageH.intValue() - borderHigh - topMarH - bottomMarH - headerMargin - footerMargin - titleHeight;
    }

    /**
     * 设置自适应一页
     */
    private void setAutoFit(BigTable bigTable, Table table, WordExportContext ctx) {
        if (bigTable.getAutoFit()) {
            // 有效行数
            int tableRows = table.getRows().size();
            // 1000为冗余高度
            long effectiveHeight = getEffectiveHeight(bigTable, ctx);
            if (effectiveHeight < 0) {
                return;
            }
            // 单行行高
            long divideHigh = effectiveHeight / tableRows;
            // todo 有问题,设置行高后不自适应
//            setRowHigh(table, divideHigh);
        }
    }

    /**
     * 设置行高
     */
    private void setRowHigh(Table table, Long rowHigh) {
        for (TableRow row : table.getRows()) {
            row.setHeight(Math.toIntExact(rowHigh));
        }
    }

}
