package com.demo;

public class Main {
    public static void main(String[] args) {

/*
        //docx加页眉
        try {
            DocxOperation.docxAddHeader("F:\\a.docx","F:\\b.docx", "合同编号：0000000000000###########000012");
        } catch (Exception ex) {
            ex.printStackTrace();
            return;
        }
*/

/*
        //docx加水印
        try {
            DocxOperation.docxAddWatermark("F:\\x.docx","F:\\xx.docx", "合同编号：888888888###888888");
        } catch (Exception ex) {
            ex.printStackTrace();
            return;
        }
*/

/*
        //aspose docx加水印
        try {
            AsposeOperation.docxAddWatermark("F:\\w.doc","F:\\xx.docx", "合同编号：888888888###888888");
        } catch (Exception ex) {
            ex.printStackTrace();
            return;
        }
*/


        //aspose docx加页眉
        try {
            AsposeOperation.docxAddHeader("F:\\6.doc","F:\\xx.docx", "合同编号：0000000000000###########000012");
        } catch (Exception ex) {
            ex.printStackTrace();
            return;
        }

/*

        //pdf加水印
        try {
            PdfOperation.pdfAddWatermark("f:/1.pdf","f:/2.pdf", "合同编号：888888888###888888");
        } catch (Exception ex) {
            ex.printStackTrace();
            return;
        }
*/

    }

}
