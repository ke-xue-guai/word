package com.export.word;//package com.export.word;
//
//import com.export.word.style.FontStyle;
//import org.docx4j.jaxb.Context;
//import org.docx4j.openpackaging.exceptions.Docx4JException;
//import org.docx4j.openpackaging.exceptions.InvalidFormatException;
//import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
//import org.docx4j.openpackaging.parts.PartName;
//import org.docx4j.openpackaging.parts.WordprocessingML.FooterPart;
//import org.docx4j.relationships.Relationship;
//import org.docx4j.wml.*;
//
//import java.math.BigInteger;
//import java.util.List;
//
//public class Footer {
//
//    public FooterPart toElement(WordExportContext ctx) {
//        String endStr = "";
//        if (ctx.footerPage != 1) {
//            endStr = String.valueOf(ctx.footerPage);
//        }
//        FooterPart footerPart = null;
//        try {
//            footerPart = new FooterPart();
//            footerPart.setJaxbElement(createFooter(ctx));
//        } catch (Docx4JException e) {
//            e.printStackTrace();
//        }
//        try {
//            // 设置名称
//            footerPart.setPartName(new PartName("/word/footer" + endStr + ".xml"));
//        } catch (InvalidFormatException e) {
//            e.printStackTrace();
//        }
//        return footerPart;
//    }
//
//    /**
//     * 创建页眉内容
//     */
//    private Ftr createFooter(WordExportContext ctx) throws Docx4JException {
//        Ftr ftr = new Ftr();
//        FontStyle fontStyle = new FontStyle();
//        fontStyle.setFontName("微软雅黑");
//        fontStyle.setFontSize(15);
//        // 页脚
//        String value = "- " + ctx.footerPage + " -";
//        Paragraph paragraph1 = new Paragraph(value, new TextAlign(TextAlign.Type.CENTER));
//        paragraph1.toElement(ctx, ftr.getContent());
//        return ftr;
//    }
//
//    public static int getHeight() {
//        // 页脚所占默认设为600
//        return 600;
//    }
//
//    public void reference(WordprocessingMLPackage wordMLPackage, WordExportContext ctx) {
//        try {
//            FooterPart footerPart = this.toElement(ctx);
//            Relationship relationship = wordMLPackage.getMainDocumentPart().addTargetPart(footerPart);
//            SectPr sectPr = getSect(wordMLPackage, ctx);
//            FooterReference footerReference = Context.getWmlObjectFactory().createFooterReference();
//            footerReference.setId(relationship.getId());
//            footerReference.setType(HdrFtrRef.DEFAULT);
//            sectPr.getEGHdrFtrReferences().add(footerReference);
//        } catch (Docx4JException e) {
//            e.printStackTrace();
//        }
//    }
//
//    /**
//     * 获取属性设置页眉页脚上下边距
//     */
//    private SectPr getSect(WordprocessingMLPackage wordMLPackage, WordExportContext ctx) {
//        SectPr sourceSectPr = wordMLPackage.getDocumentModel().getSections().get(0).getSectPr();
//        List<Object> content = wordMLPackage.getMainDocumentPart().getContent();
//        // 创建段
//        P p = new P();
//        PPr pPr = new PPr();
//        p.setPPr(pPr);
//        // 设置当前样式
//        SectPr currentSectPr = Context.getWmlObjectFactory().createSectPr();
//        currentSectPr.setPgSz(sourceSectPr.getPgSz());
//        currentSectPr.setPgMar(sourceSectPr.getPgMar());
//        if (ctx.footerPage != 1) {
//            SectPr.Type type = new SectPr.Type();
//            type.setVal("continuous");
//            currentSectPr.setType(type);
//        }
//        CTDocGrid ctDocGrid = new CTDocGrid();
//        ctDocGrid.setType(STDocGrid.LINES);
//        ctDocGrid.setLinePitch(BigInteger.valueOf(312));
//        ctDocGrid.setCharSpace(BigInteger.ZERO);
//        currentSectPr.setDocGrid(ctDocGrid);
//        CTColumns ctColumns = new CTColumns();
//        ctColumns.setNum(BigInteger.ONE);
//        ctColumns.setSpace(BigInteger.valueOf(425));
//        currentSectPr.setCols(ctColumns);
//        // 设置当前样式
//        pPr.setSectPr(currentSectPr);
//
//        content.add(p);
//        return currentSectPr;
//    }
//}
