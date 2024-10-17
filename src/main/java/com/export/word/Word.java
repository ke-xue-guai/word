package com.export.word;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Body;
import org.docx4j.wml.SectPr;

import java.io.ByteArrayOutputStream;
import java.util.ArrayList;
import java.util.List;


public class Word {


    private List<Element> elements = new ArrayList<>();

    /**
     * 页面样式
     */
    private WordStyle wordStyle;

    public void addElement(Element element) {
        this.elements.add(element);
    }


    public ByteArrayOutputStream toOutStream() {
        WordExportContext ctx = new WordExportContext();
        ByteArrayOutputStream rst = new ByteArrayOutputStream();
        WordprocessingMLPackage wordPackage = null;
        List<Object> content = null;
        try {
            wordPackage = WordprocessingMLPackage.createPackage();
            MainDocumentPart mainDocumentPart = wordPackage.getMainDocumentPart();
            content = mainDocumentPart.getContents().getContent();
            Body body = mainDocumentPart.getContents().getBody();
            if (this.wordStyle != null) {
                body.setSectPr(wordStyle.get());
                ctx.wordMargins = this.wordStyle.getMargins();
                SectPr.PgSz pgSz = body.getSectPr().getPgSz();
                ctx.pageHeight = pgSz.getH();
                ctx.pageWidth = pgSz.getW();
//                ctx.isFixRowHigh = wordStyle.getFixRowHigh();
            }
        } catch (Docx4JException e) {
            e.printStackTrace();
        }

        for (Element element : this.elements) {
            element.toElement(ctx, content);
        }
        try {
            wordPackage.save(rst);
        } catch (Docx4JException e) {
            e.printStackTrace();
        }
        return rst;
    }

    public List<Element> getElements() {
        return elements;
    }

    public void setElements(List<Element> elements) {
        this.elements = elements;
    }

    public WordStyle getWordStyle() {
        return wordStyle;
    }

    public void setWordStyle(WordStyle wordStyle) {
        this.wordStyle = wordStyle;
    }

}
