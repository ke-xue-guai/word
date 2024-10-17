package com.export.word.table;

import com.export.word.Border;
import com.export.word.Single;
import org.docx4j.wml.TcPrInner;

/**
 * 单元格的边框
 */
public class TableColBorders extends Single<TcPrInner.TcBorders> {

    private Border top;
    private Border left;
    private Border right;
    private Border bottom;

    public TableColBorders() {

    }

    /**
     * 上下左右为同一个样式
     */
    public TableColBorders(Border one) {
        this.top = one;
        this.left = one;
        this.right = one;
        this.bottom = one;
    }

    /**
     * 上下为同一个样式
     * 左右为同一个样式
     */
    public TableColBorders(Border one, Border two) {
        this.top = one;
        this.left = two;
        this.right = two;
        this.bottom = one;
    }

    public TableColBorders(Border top, Border left, Border right, Border bottom) {
        this.top = top;
        this.left = left;
        this.right = right;
        this.bottom = bottom;
    }

    public Border getTop() {
        return top;
    }

    public void setTop(Border top) {
        this.top = top;
    }

    public Border getLeft() {
        return left;
    }

    public void setLeft(Border left) {
        this.left = left;
    }

    public Border getRight() {
        return right;
    }

    public void setRight(Border right) {
        this.right = right;
    }

    public Border getBottom() {
        return bottom;
    }

    public void setBottom(Border bottom) {
        this.bottom = bottom;
    }

    @Override
    public TcPrInner.TcBorders create() {
        TcPrInner.TcBorders borders = new TcPrInner.TcBorders();
        if (this.top != null) {
            borders.setTop(this.top.get());
        }
        if (this.left != null) {
            borders.setLeft(this.left.get());
        }
        if (this.right != null) {
            borders.setRight(this.right.get());
        }
        if (this.bottom != null) {
            borders.setBottom(this.bottom.get());
            borders.setInsideH(this.bottom.get());
            borders.setInsideV(this.bottom.get());
        }
        return borders;
    }

}
