package com.demo;

import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.xmlbeans.impl.xb.xmlschema.SpaceAttribute;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHdrFtr;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;

public class DocxOperation {

    public static void docxAddWatermark(String srcFile, String WatermarkString) throws Exception {
        docxAddWatermark(srcFile, srcFile, WatermarkString);
    }

    public static void docxAddWatermark(String srcFile, String descFile, String WatermarkString) throws Exception {
        File is = new File(srcFile);//文件路径
        FileInputStream fis = new FileInputStream(is);
        XWPFDocument docx = new XWPFDocument(fis);
        MyXWPFHeaderFooterPolicy footer = new MyXWPFHeaderFooterPolicy(docx);
        footer.createWatermark(WatermarkString);

        fis.close();

        FileOutputStream out = new FileOutputStream(descFile);
        docx.write(out);

        docx.close();
        out.close();
    }


    public static void docxAddHeader(String srcFile, String headerString) throws Exception
    {
        docxAddHeader(srcFile, srcFile, headerString);
    }


    public static void docxAddHeader(String srcFile, String descFile, String headerString) throws Exception {
        File is = new File(srcFile);//文件路径
        FileInputStream fis = new FileInputStream(is);
        XWPFDocument docx = new XWPFDocument(fis);//文档对象

        CTP ctp = CTP.Factory.newInstance();
        XWPFParagraph paragraph = new XWPFParagraph(ctp, docx);//段落对象
        ctp.addNewR().addNewT().setStringValue(headerString);//设置页眉参数
        ctp.addNewR().addNewT().setSpace(SpaceAttribute.Space.PRESERVE);

        CTSectPr sectPr = docx.getDocument().getBody().isSetSectPr() ? docx.getDocument().getBody().getSectPr() : docx.getDocument().getBody().addNewSectPr();
        XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(docx, sectPr);


        XWPFHeader header = policy.getHeader(XWPFHeaderFooterPolicy.DEFAULT);
        if (header == null)
        {
            header = policy.createHeader(STHdrFtr.DEFAULT, new XWPFParagraph[] { paragraph });
        } else {
            header.createParagraph().createRun().setText(headerString);
        }

        header.setXWPFDocument(docx);
        fis.close();

        OutputStream out = new FileOutputStream(descFile);
        docx.write(out);//输出到本地
        docx.close();

        out.close();

    }

}
