package priv.viking.word;

import org.apache.poi.POIXMLDocument;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.io.IOException;
import java.util.List;

/**
 * Created by Viking on 2019/4/3
 * 读取word文档中的内容
 */
public class ReadWord {
    public static void main(String[] args) throws IOException {
        String path = "F:/poi-word-test/ReadDoc.docx";
        XWPFDocument document = new XWPFDocument(POIXMLDocument.openPackage(path));
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        for (XWPFParagraph paragraph : paragraphs){
            System.out.println(paragraph.getText());
//            List<XWPFRun> runs = paragraph.getRuns();
//            for (XWPFRun run : runs){
//                System.out.println(run.text());
//            }
        }
    }
}
