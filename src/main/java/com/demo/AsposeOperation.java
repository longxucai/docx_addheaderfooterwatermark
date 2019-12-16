package com.demo;


import com.aspose.words.Shape;
import com.aspose.words.*;

import java.awt.*;
import java.io.FileOutputStream;

public class AsposeOperation {


    public static void docxAddWatermark(String srcFile, String WatermarkString) throws Exception {
        docxAddWatermark(srcFile, srcFile, WatermarkString);
    }

    public static void docxAddWatermark(String srcFile, String descFile, String WatermarkString) throws Exception {

        Document doc = new Document(srcFile);

        Shape watermark = new Shape(doc, ShapeType.TEXT_PLAIN_TEXT);
        //水印内容
        watermark.getTextPath().setText(WatermarkString);
        //水印字体
        watermark.getTextPath().setFontFamily("宋体");
        //水印宽度
        watermark.setWidth(500);
        //水印高度
        watermark.setHeight(100);
        //旋转水印
        watermark.setRotation(-40);
        //水印颜色
        watermark.getFill().setColor(Color.lightGray);
        watermark.setStrokeColor(Color.lightGray);
        watermark.setRelativeHorizontalPosition(RelativeHorizontalPosition.PAGE);
        watermark.setRelativeVerticalPosition(RelativeVerticalPosition.PAGE);
        watermark.setWrapType(WrapType.NONE);
        watermark.setVerticalAlignment(VerticalAlignment.CENTER);
        watermark.setHorizontalAlignment(HorizontalAlignment.CENTER);
        Paragraph watermarkPara = new Paragraph(doc);
        watermarkPara.appendChild(watermark);
        for (Section sect : doc.getSections())
        {
            insertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HEADER_PRIMARY);
            insertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HEADER_FIRST);
            insertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HEADER_EVEN);
        }

        FileOutputStream out = new FileOutputStream(descFile);
        doc.save(out, SaveFormat.DOCX);

        out.close();
    }


    private static void insertWatermarkIntoHeader(Paragraph watermarkPara, Section sect, int headerType) throws Exception{
        HeaderFooter header = sect.getHeadersFooters().getByHeaderFooterType(headerType);
        if (header == null)
        {
            header = new HeaderFooter(sect.getDocument(), headerType);
            sect.getHeadersFooters().add(header);
        }
        header.appendChild(watermarkPara.deepClone(true));
    }





    public static void docxAddHeader(String srcFile, String headerString) throws Exception
    {
        docxAddHeader(srcFile, srcFile, headerString);
    }


    public static void docxAddHeader(String srcFile, String descFile, String headerString) throws Exception {
        Document doc = new Document(srcFile);

/*
        for (Section sect : doc.getSections())
        {
            HeaderFooter header;
            header = sect.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.HEADER_PRIMARY);
            if (header != null) {
                header.removeAllChildren();
            }

            header = sect.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.HEADER_FIRST);
            if (header != null) {
                header.removeAllChildren();
            }

            header = sect.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.HEADER_EVEN);
            if (header != null) {
                header.removeAllChildren();
            }
        }*/

        for (Section sect : doc.getSections())
        {
            HeaderFooter header;
            boolean isExistsHeader;

            header = sect.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.HEADER_PRIMARY);
            isExistsHeader = false;
            if (header == null)
            {
                header = new HeaderFooter(sect.getDocument(), HeaderFooterType.HEADER_PRIMARY);
                isExistsHeader = true;
            }
            header.appendParagraph(headerString);
            if (isExistsHeader)
                sect.getHeadersFooters().add(header);



            header = sect.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.HEADER_FIRST);
            isExistsHeader = false;
            if (header == null)
            {
                header = new HeaderFooter(sect.getDocument(), HeaderFooterType.HEADER_FIRST);
                isExistsHeader = true;
            }
            header.appendParagraph(headerString);
            if (isExistsHeader)
                sect.getHeadersFooters().add(header);



            header = sect.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.HEADER_EVEN);
            isExistsHeader = false;
            if (header == null)
            {
                header = new HeaderFooter(sect.getDocument(), HeaderFooterType.HEADER_EVEN);
                isExistsHeader = true;
            }
            header.appendParagraph(headerString);
            if (isExistsHeader)
                sect.getHeadersFooters().add(header);



        }

        FileOutputStream out = new FileOutputStream(descFile);
        doc.save(out, SaveFormat.DOCX);

        out.close();

    }






}
