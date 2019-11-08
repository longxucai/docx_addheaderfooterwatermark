package com.demo;

public class Main {
    public static void main(String[] args) {

/*
        //docx加页眉
        try {
            DocxOperation.docxAddHeader("F:\\创梦广告平台合作协议（开发者）_V1_SL0170829_clean_new.docx","F:\\b.docx", "合同编号：0000000000000###########000011");
        } catch (Exception ex) {
            ex.printStackTrace();
            return;
        }
*/


        //docx加水印
        try {
            DocxOperation.docxAddWatermark("F:\\a.docx","F:\\b.docx", "合同编号：888888888###888888");
        } catch (Exception ex) {
            ex.printStackTrace();
            return;
        }



        //pdf加水印
        try {
            PdfOperation.pdfAddWatermark("f:/1.pdf","f:/2.pdf", "合同编号：888888888###888888");
        } catch (Exception ex) {
            ex.printStackTrace();
            return;
        }


    }

}
