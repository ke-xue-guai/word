package com.export.word.style;


public class PageStyle {

    /**
     * 横版、竖版
     */
    public static final short LANDSCAPE = 0;
    public static final short PORTRAIT = 1;

    public static final PageStyle DEFAULT = new PageStyle();

    static {
        DEFAULT.setPageH(11907);
        DEFAULT.setPageW(19839);
        DEFAULT.setLeftMar(415);
        DEFAULT.setRightMar(415);
        DEFAULT.setTopMar(200);
        DEFAULT.setBottomMar(200);
        DEFAULT.setHeader(400);
        DEFAULT.setFooter(200);
    }

    /**
     * 纸张宽度
     */
    private Integer pageW;

    /**
     * 纸张高度
     */
    private Integer pageH;

    /**
     * 摆放方式
     */
    private short orientation;

    /**
     * 左页边距
     */
    private Integer leftMar;

    /**
     * 右页边距
     */
    private Integer rightMar;

    /**
     * 下页边距
     */
    private Integer bottomMar;

    /**
     * 上页边距
     */
    private Integer topMar;

    /**
     * 页脚顶端距离
     */
    private Integer footer;

    /**
     * 页眉顶端距离
     */
    private Integer header;


    public Integer getPageW() {
        return pageW;
    }

    public void setPageW(Integer pageW) {
        this.pageW = pageW;
    }

    public Integer getPageH() {
        return pageH;
    }

    public void setPageH(Integer pageH) {
        this.pageH = pageH;
    }

    public short getOrientation() {
        return orientation;
    }

    public void setOrientation(short orientation) {
        this.orientation = orientation;
    }

    public Integer getLeftMar() {
        return leftMar;
    }

    public void setLeftMar(Integer leftMar) {
        this.leftMar = leftMar;
    }

    public Integer getRightMar() {
        return rightMar;
    }

    public void setRightMar(Integer rightMar) {
        this.rightMar = rightMar;
    }

    public Integer getBottomMar() {
        return bottomMar;
    }

    public void setBottomMar(Integer bottomMar) {
        this.bottomMar = bottomMar;
    }

    public Integer getTopMar() {
        return topMar;
    }

    public void setTopMar(Integer topMar) {
        this.topMar = topMar;
    }

    public Integer getFooter() {
        return footer;
    }

    public void setFooter(Integer footer) {
        this.footer = footer;
    }

    public Integer getHeader() {
        return header;
    }

    public void setHeader(Integer header) {
        this.header = header;
    }
}
