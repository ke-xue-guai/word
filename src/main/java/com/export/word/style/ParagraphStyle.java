package com.export.word.style;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class ParagraphStyle {


    /**
     * 段摆放格式（居中、居左、居右）
     */
    private Alignment alignment;

    public enum Alignment {
        PARAGRAPH_LEFT, PARAGRAPH_CENTER, PARAGRAPH_RIGHT
    }

}
