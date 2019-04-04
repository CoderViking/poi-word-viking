import org.apache.poi.POIXMLDocument;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by yanshuai on 2019/4/3
 * 段落替换
 */
public class ReplaceGraph {
    public static void main(String[] args) throws IOException {
        String path = "F:/poi-word-test/ReadDoc.docx";
        XWPFDocument document = new XWPFDocument(POIXMLDocument.openPackage(path));
        Map<String,String> map = new HashMap<>();
        List<String> list = new ArrayList<>();

    }
    //替换段落文本
    public void changeText(XWPFDocument document, Map<String,String> textMap){
        //获取段落集合
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        int i = 0;
        for(XWPFParagraph paragraph : paragraphs) {
            System.out.println("i--->"+i++);
//            replaceParagraph(paragraph,textMap);
        }
    }
}
