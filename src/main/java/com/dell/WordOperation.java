package com.dell;

import com.dell.util.Tool;
import org.apache.poi.POIXMLDocument;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.xmlbeans.XmlException;
import org.junit.Test;

import java.io.File;
import java.io.IOException;

public class WordOperation {
    public static void traverFolder(File file,String key) throws OpenXML4JException, XmlException, IOException {
        File[] files = file.listFiles();
        if (files!=null){
            for (File file1 : files) {
                if (file1.isDirectory()) {
                    traverFolder(file1,key);
                }else {
                    if (file1.toString().endsWith(".docx")){
                        Tool.findKeyWords(file1.toString(),key);
                    }
                }
            }
        }
    }
    @Test
    public void test(/**/) throws IOException, InterruptedException, OpenXML4JException, XmlException {
//        String path = "C:\\Users\\15756\\Desktop\\一班_第十周考试\\第八组\\后端\\李佳澳_后端.docx";
        String path = "C:\\Users\\15756\\Desktop\\一班_第十周考试";
        traverFolder(new File(path),"16");
    }
}
