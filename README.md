# docx_addheaderfooterwatermark
给word(docx)文档添加页眉、页脚、水印

## 背景
一个已存在的合同管理项目，需要在上传word文档之后，在文档中自动加上合同编号。

## 方案
通过POI对word文档进行操作，发现如果文档已经存在页眉和页脚，加水印的功能无效。经过调试，POI的org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy类貌似有bug，临时重写该类(com.demo.MyXWPFHeaderFooterPolicy)解决。

主要修改点：
```
        public XWPFHeader createHeader(STHdrFtr.Enum type, XWPFParagraph[] pars) {
            XWPFHeader header = this.getHeader(type);
            if (header == null) {
                HdrDocument hdrDoc = HdrDocument.Factory.newInstance();
                XWPFRelation relation = XWPFRelation.HEADER;
                int i = this.getRelationIndex(relation);
                XWPFHeader wrapper = (XWPFHeader)this.doc.createRelationship(relation, XWPFFactory.getInstance(), i);
                wrapper.setXWPFDocument(this.doc);
                CTHdrFtr hdr = this.buildHdr(type, wrapper, pars);
                wrapper.setHeaderFooter(hdr);
                hdrDoc.setHdr(hdr);
                this.assignHeader(wrapper, type);
                //header = wrapper;
    
                return wrapper;
            }
    
            //header已存在时
            HdrDocument hdrDoc = HdrDocument.Factory.newInstance();
            XWPFRelation relation = XWPFRelation.HEADER;
            int i = this.getRelationIndex(relation);
            XWPFHeader wrapper = (XWPFHeader)this.doc.createRelationship(relation, XWPFFactory.getInstance(), i);
            header.setXWPFDocument(this.doc);
            CTHdrFtr hdr = this.buildHdr(type, header, pars);
            header.setHeaderFooter(hdr);
            hdrDoc.setHdr(hdr);
            this.assignHeader(header, type);
            //header = wrapper;
    
            return header;
        }
        
    private CTHdrFtr buildHdrFtr(XWPFParagraph[] paragraphs, XWPFHeaderFooter wrapper) {
        CTHdrFtr ftr = wrapper._getHdrFtr();
        int n = ftr.getPArray().length;

        if (paragraphs != null) {
            for(int i = 0; i < paragraphs.length; ++i) {
                ftr.addNewP();
                //ftr.setPArray(i, paragraphs[i].getCTP());
                //赋值给新加的元素
                ftr.setPArray(n+i, paragraphs[i].getCTP());
            }
        }

        return ftr;
    }
    
```


另外，关于水印样式也做了调整
```
        //shape.setStyle("position:absolute;margin-left:0;margin-top:0;width:415pt;height:207.5pt;z-index:-251654144;mso-wrap-edited:f;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin");
        shape.setStyle("position:absolute;margin-left:0;margin-top:0;width:415pt;height:207.5pt;z-index:-251654144;mso-wrap-edited:f;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin;rotation:315;height:30pt");
```

## 参考
https://github.com/ymyang/watermark
