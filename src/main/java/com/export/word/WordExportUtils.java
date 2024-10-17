package com.export.word;

import com.export.word.style.Cell;
import com.export.word.style.ParagraphStyle;
import com.export.word.style.Word;
import org.docx4j.jaxb.Context;
import org.docx4j.model.properties.table.tr.TrHeight;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;


import java.math.BigInteger;
import java.util.List;
import java.util.Optional;

public class WordExportUtils extends NewWordUtils {

    /**
     * docx4j合并word文档
     */
    public static WordprocessingMLPackage docx4jMergeWord(List<WordprocessingMLPackage> wordProcessingMLPackageList) {
        WordprocessingMLPackage wordMLPackage = null;
        try {
            // 创建一个新的Word文档
            wordMLPackage = WordprocessingMLPackage.createPackage();
            // 将内容从每个文档复制到合并的文档中
            for (WordprocessingMLPackage wordprocessingMLPackage : wordProcessingMLPackageList) {
                wordMLPackage.getMainDocumentPart().getContent().addAll(wordprocessingMLPackage.getMainDocumentPart().getContent());
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return wordMLPackage;
    }

    /**
     * docx4j生成word文档
     */
    public static WordprocessingMLPackage docx4jCreatWord(Word word) {
        WordprocessingMLPackage wordMLPackage = null;
        try {
            //创建一个新的Word文档对象
            wordMLPackage = WordprocessingMLPackage.createPackage();
            //创建一个新的内容对象
            MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();
            ObjectFactory factory = Context.getWmlObjectFactory();
            // 创建一个节（Section）
            SectPr sectionProperties = factory.createSectPr();
            // 设置纸张样式
            creatPageStyle(sectionProperties);
            // 设置页边距
            creatPageMargins(sectionProperties);

//            // 添加页脚
//            FooterPart footerPart = new FooterPart();
//            wordMLPackage.getMainDocumentPart().addTargetPart(footerPart);
//            // 页脚内容
//            Ftr footer = factory.createFtr();
//            // 段落
//            P p = factory.createP();
//            // 文本运行
//            R r = factory.createR();
//            // 文本内容
//            Text factoryText = factory.createText();
//            factoryText.setValue("页脚内容");
//            r.getContent().add(factoryText);
//            p.getContent().add(r);
//            footer.getContent().add(p);
//            footerPart.setJaxbElement(footer);
//            // 将页脚引用与页脚部分关联
//            sectionProperties.getEGHdrFtrReferences().add(factory.createHeaderReference());

            mainDocumentPart.getContents().getBody().setSectPr(sectionProperties);
            // 单元格数据
            List<Cell> cellList = word.getCellList();
            //创建一个表格对象
            Tbl table = factory.createTbl();
            // 创建表格行和单元格
            for (int i = 0; i < word.getRows(); i++) {
                Tr tableRow = factory.createTr();
                TrPr trPr = tableRow.getTrPr() == null ? factory.createTrPr() : tableRow.getTrPr();
                CTHeight ctHeight = new CTHeight();
                ctHeight.setVal(BigInteger.valueOf(300));
                TrHeight trHeight = new TrHeight(ctHeight);
                trHeight.set(trPr);
                tableRow.setTrPr(trPr);

                int finalI = i;
                for (int j = 0; j < word.getCols(); j++) {
                    int finalJ = j;
                    Optional<Cell> cellOptional = cellList.stream()
                            .filter(it -> it.getPositionX() == finalI && it.getPositionY() == finalJ)
                            .findAny();
                    if (!cellOptional.isPresent()) {
                        continue;
                    }
                    Cell cell = cellOptional.get();
                    Tc tableCell = factory.createTc();
                    // 合并列
                    if (cell.getRowspan() != 1 || cell.getColspan() != 1) {
                        // 单元格，合并到下一列
                        TcPr tcpr = factory.createTcPr();
                        TcPrInner.GridSpan gridSpan = factory.createTcPrInnerGridSpan();
                        // 设置合并的列数
                        gridSpan.setVal(BigInteger.valueOf(cell.getColspan()));
                        tcpr.setGridSpan(gridSpan);
                        tableCell.setTcPr(tcpr);
                    }
                    // 设置边框
                    TcPr tcPr = tableCell.getTcPr() == null ? factory.createTcPr() : tableCell.getTcPr();
                    TcPrInner.TcBorders tcBorders = tcPr.getTcBorders() == null ? factory.createTcPrInnerTcBorders() : tcPr.getTcBorders();
                    org.docx4j.wml.CTBorder border = factory.createCTBorder();
                    // 设置边框颜色
                    border.setColor("auto");
                    // 设置边框大小
                    border.setSz(BigInteger.valueOf(4));
                    // 设置边框间距
                    border.setSpace(BigInteger.valueOf(0));
                    if (i >= 3) {
                        // 设置边框样式
                        border.setVal(org.docx4j.wml.STBorder.SINGLE);
                    } else {
                        // 设置边框样式
                        border.setVal(org.docx4j.wml.STBorder.NONE);
                    }
//                    border.setVal(org.docx4j.wml.STBorder.SINGLE);
                    tcBorders.setTop(border);
                    tcBorders.setBottom(border);
                    tcBorders.setLeft(border);
                    tcBorders.setRight(border);
                    tcPr.setTcBorders(tcBorders);

                    //设置单元格边距
                    // 左边距
                    TcMar cellMargin = new TcMar();
                    TblWidth leftMargin = new TblWidth();
                    leftMargin.setW(BigInteger.valueOf(50));
                    cellMargin.setLeft(leftMargin);

                    TblWidth rightMargin = new TblWidth();
                    rightMargin.setW(BigInteger.valueOf(50));
                    cellMargin.setRight(rightMargin);

                    TblWidth topMargin = new TblWidth();
                    topMargin.setW(BigInteger.valueOf(30));
                    cellMargin.setTop(topMargin);

                    TblWidth bottomMargin = new TblWidth();
                    bottomMargin.setW(BigInteger.valueOf(30));
                    cellMargin.setBottom(bottomMargin);

                    tcPr.setTcMar(cellMargin);

                    // 设置垂直对齐方式
                    CTVerticalJc verticalJc = factory.createCTVerticalJc();
                    verticalJc.setVal(STVerticalJc.CENTER);
                    // 将垂直对齐方式应用到单元格样式对象
                    tcPr.setVAlign(verticalJc);

//                    // 设置单元格背景色
//                    // 创建填充对象
//                    CTShd fill = new CTShd();
//                    fill.setFill(cell.getCellStyle().getBgColor());
//                    tcPr.setShd(fill);

                    tableCell.setTcPr(tcPr);

                    P cellParagraph = factory.createP();
                    // 设置单元格水平对齐方式
                    Jc justification = new Jc();
                    ParagraphStyle.Alignment alignment = cell.getParagraphStyle().getAlignment();
                    if (alignment != null) {
                        switch (alignment) {
                            case PARAGRAPH_LEFT: {
                                justification.setVal(JcEnumeration.LEFT);
                                break;
                            }
                            case PARAGRAPH_CENTER: {
                                justification.setVal(JcEnumeration.CENTER);
                                break;
                            }
                            case PARAGRAPH_RIGHT: {
                                justification.setVal(JcEnumeration.RIGHT);
                                break;
                            }
                            default:
                        }
                    }
                    PPr paragraphProperties = cellParagraph.getPPr();
                    if (paragraphProperties == null) {
                        paragraphProperties = new PPr();
                        cellParagraph.setPPr(paragraphProperties);
                    }
                    paragraphProperties.setJc(justification);

                    R run = factory.createR();
                    Text text = factory.createText();
                    text.setValue(cell.getValue() == null ? "" : cell.getValue());
                    run.getContent().add(text);
                    cellParagraph.getContent().add(run);
                    tableCell.getContent().add(cellParagraph);

                    // 设置字体
                    RPr runProperties = run.getRPr() == null ? factory.createRPr() : run.getRPr();

                    HpsMeasure fontSize = factory.createHpsMeasure();
                    // 设置字体大小
                    fontSize.setVal(BigInteger.valueOf(cell.getFontStyle().getFontSize()));
                    runProperties.setSz(fontSize);

                    RFonts runFonts = factory.createRFonts();
                    runFonts.setAscii("Calibri");
                    // 设置中文字体
                    runFonts.setHAnsi("SimSun");
                    runProperties.setRFonts(runFonts);

                    // 创建字体加粗标记对象,必须在设置中文字体之后执行才能使中文生效
                    BooleanDefaultTrue bold = new BooleanDefaultTrue();
                    bold.setVal(cell.getFontStyle().isBold());
                    runProperties.setB(bold);

                    run.setRPr(runProperties);

                    tableRow.getContent().add(tableCell);
                }
                table.getContent().add(tableRow);
            }

            mainDocumentPart.getContent().add(table);

//            FooterPart footerPart = new FooterPart();
//            footerPart.setPackage(wordMLPackage);
//
//            // 生成页码
//            Ftr ftr = factory.createFtr();
//            P paragraph = factory.createP();
//
//            // 创建开始字段
//            {
//                R run = factory.createR();
//                FldChar fldchar = factory.createFldChar();
//                fldchar.setFldCharType(org.docx4j.wml.STFldCharType.BEGIN);
//                run.getContent().add(fldchar);
//                paragraph.getContent().add(run);
//            }
//
//            // 创建页码
//            {
//                R run = factory.createR();
//                Text txt = new Text();
//                txt.setValue("321页脚");
////                txt.setSpace("preserve");
////                txt.setValue("NUMPAGES ");
////                txt.setValue(" PAGE   \\* MERGEFORMAT ");
//                run.getContent().add(factory.createRInstrText(txt));
//                paragraph.getContent().add(run);
//            }
//
//            // 创建结束字段
//            {
//                R run = factory.createR();
//                FldChar fldchar = factory.createFldChar();
//                fldchar.setFldCharType(org.docx4j.wml.STFldCharType.END);
//                run.getContent().add(fldchar);
//                paragraph.getContent().add(run);
//            }
//            // 设置页脚高度
//            PPr pPr = factory.createPPr();
//            PPrBase.Spacing spacing = new PPrBase.Spacing();
//            spacing.setAfter(BigInteger.valueOf(100));
//            pPr.setSpacing(spacing);
//            paragraph.setPPr(pPr);
//
//            ftr.getContent().add(paragraph);
//            footerPart.setJaxbElement(ftr);
//
//
//            Relationship relationship = mainDocumentPart.addTargetPart(footerPart);
//
//            List<SectionWrapper> sections = wordMLPackage.getDocumentModel().getSections();
//
//            SectPr sectionPrs = sections.get(sections.size() - 1).getSectPr();
//            if (sectionPrs == null) {
//                sectionPrs = factory.createSectPr();
//                wordMLPackage.getMainDocumentPart().addObject(sectionPrs);
//                sections.get(sections.size() - 1).setSectPr(sectionPrs);
//            }
//
//            FooterReference footerReference = factory.createFooterReference();
//            footerReference.setId(relationship.getId());
//            footerReference.setType(HdrFtrRef.DEFAULT);
//            sectionPrs.getEGHdrFtrReferences().add(footerReference);

        } catch (Exception e) {
            e.printStackTrace();
        }
        return wordMLPackage;
    }

    /**
     * 设置纸张样式
     */
    private static void creatPageStyle(SectPr sectionProperties) {
        ObjectFactory factory = Context.getWmlObjectFactory();
        // 创建纸张样式
        SectPr.PgSz pageSize = factory.createSectPrPgSz();
        // 设置宽度
        pageSize.setW(BigInteger.valueOf(15840));
        // 设置高度
        pageSize.setH(BigInteger.valueOf(12240));
        // 创建页面方向对象并设置为横向
        pageSize.setOrient(org.docx4j.wml.STPageOrientation.LANDSCAPE);

        // 将纸张样式应用于节
        sectionProperties.setPgSz(pageSize);
    }

    /**
     * 设置页面边距
     */
    private static void creatPageMargins(SectPr sectionProperties) {
        ObjectFactory factory = Context.getWmlObjectFactory();
        // 创建一个页边距
        SectPr.PgMar pageMargins = factory.createSectPrPgMar();
        // 设置左边距 -- 单位：TWIP = 磅*10
        pageMargins.setLeft(BigInteger.valueOf(283));
        // 设置右边距
        pageMargins.setRight(BigInteger.valueOf(283));
        // 设置上边距
        pageMargins.setTop(BigInteger.valueOf(425));
        // 设置下边距
        pageMargins.setBottom(BigInteger.valueOf(454));
        // 将页边距应用于节
        sectionProperties.setPgMar(pageMargins);
    }


}
