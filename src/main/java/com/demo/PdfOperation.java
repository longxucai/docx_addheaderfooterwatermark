package com.demo;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Element;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.*;

import java.io.FileOutputStream;

public class PdfOperation {

    public static void pdfAddWatermark(String srcFile, String descFile, String WatermarkString) throws Exception {
        PdfReader reader = null;
        PdfStamper pdfStamper = null;
        try {
            reader = new PdfReader(srcFile);
            pdfStamper = new PdfStamper(reader, new FileOutputStream(descFile));

            addWatermark(pdfStamper, WatermarkString);
        } finally {
            if (pdfStamper != null) {
                pdfStamper.close();
            }
        }
    }

    private static void addWatermark(PdfStamper pdfStamper, String watermark) throws Exception {
        PdfGState gs = new PdfGState();
        // 设置透明度为0.4
        gs.setFillOpacity(0.4f);
        gs.setStrokeOpacity(0.4f);

        // 设置字体
        BaseFont base = BaseFont.createFont("STSong-Light", "UniGB-UCS2-H",BaseFont.EMBEDDED);
        int toPage = pdfStamper.getReader().getNumberOfPages();

        PdfContentByte content = null;
        Rectangle pageRect = null;
        for (int i = 1; i <= toPage; i++) {
            pageRect = pdfStamper.getReader().getPageSizeWithRotation(i);
            // 计算水印X,Y坐标
            //float x = pageRect.getWidth() / 2;
            //float y = pageRect.getHeight() / 2;
            float x = 100;
            float y = pageRect.getHeight() - 20;

            //获得PDF最顶层
            content = pdfStamper.getOverContent(i);
            content.saveState();
            // set Transparency
            content.setGState(gs);
            content.beginText();
            //content.setColorFill(BaseColor.GRAY);
            content.setColorFill(new BaseColor(85, 85, 85));
            content.setFontAndSize(base, 12);
            // 水印文字成45度角倾斜
            //content.showTextAligned(Element.ALIGN_CENTER, watermark, x, y, 315);
            content.showTextAligned(Element.ALIGN_CENTER, watermark, x, y, 0);
            content.endText();

        }
    }
}
