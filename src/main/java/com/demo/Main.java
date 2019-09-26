package com.demo;

public class Main {
    public static void main(String[] args) {

        //加页眉
        try {
            DocxOperation.docxAddHeader("F:\\1.docx","F:\\1.docx", "合同编号：0000000000000###########000011");
        } catch (Exception ex) {
            System.out.println(ex);
            return;
        }

        //加水印
        try {
            DocxOperation.docxAddWatermark("F:\\1.docx","F:\\1.docx", "合同编号：888888888###888888");
        } catch (Exception ex) {
            System.out.println(ex);
            return;
        }



    }

}
