package com.export.word;


public class NewWordUtils {

//
//    /**
//     * 多表打印
//     *
//     * @return ByteArrayOutputStream
//     */
//    public static ByteArrayOutputStream createWordProcessingMLPackage(Word word) {
//        return word.createWordProcessingMLPackage();
//    }
//
//
//    /**
//     * 多表打印
//     *
//     * @return ByteArrayOutputStream
//     */
//    public static ByteArrayOutputStream createWordNoHeader(Word word) {
//        return word.createWordNoHeader();
//    }
//
//    /**
//     * word转pdf
//     *
//     * @param word     word
//     * @param pathFile 输出路径
//     */
//    public static void toPdF(Word word, String pathFile) {
//        List<Unit> units = word.getUnits();
////        for (Unit unit : units) {
////            // 字体样式调整
////            for (AbstractElement element : unit.getElements()) {
////                if (element instanceof Table) {
////                    Table table = (Table) element;
////                    for (TableRow row : table.getRows()) {
////                        for (TableCol col : row.getCols()) {
////                            FontStyle fontStyle = col.getFontStyle();
////                            if (fontStyle == null) {
////                                col.setFontStyle(FontStyle.DEFAULT_FONT_STYLE);
////                            } else {
////                                FontStyle fontNewStyle = new FontStyle();
////                                fontNewStyle.setFontName(fontStyle.getFontName());
////                                fontNewStyle.setFontSize(fontStyle.getFontSize());
////                                fontNewStyle.setBold(fontStyle.isBold());
////                                col.setFontStyle(fontNewStyle);
////                            }
////
////                        }
////                    }
////                }
////            }
////        }
//        ByteArrayOutputStream byteArrayOutputStream = createWordProcessingMLPackage(word);
//        InputStream inputStream = new ByteArrayInputStream(byteArrayOutputStream.toByteArray());
//        WordprocessingMLPackage mlPackage = null;
//        try {
//            mlPackage = Docx4J.load(inputStream);
//        } catch (Docx4JException e) {
//            e.printStackTrace();
//        }
//        FOSettings foSettings = Docx4J.createFOSettings();
//        foSettings.setWmlPackage(mlPackage);
//        try {
//            OutputStream outputStream = new FileOutputStream(pathFile + ".pdf");
//            Docx4J.toFO(foSettings, outputStream, Docx4J.FLAG_EXPORT_PREFER_XSL);
//        } catch (Docx4JException | FileNotFoundException e) {
//            e.printStackTrace();
//        }
//    }


}
