import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlToken;
import org.apache.xmlbeans.impl.xb.xmlschema.SpaceAttribute;
import org.openxmlformats.schemas.drawingml.x2006.main.CTNonVisualDrawingProps;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPositiveSize2D;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.*;
import java.math.BigInteger;

/**
 * Created by viking on 2019/3/25
 * 创建word文档
 */
public class CreatedWord {

    public static void main(String[] args) throws Exception {
        XWPFDocument document = new XWPFDocument();
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText("第一个使用poi成功创建的word");
        run.setFontFamily("楷体");
        run.setFontSize(20);
        run.setColor("FF0000");
        run.setBold(true);
        run.setText("1测试用的文本1",1);
        run.setText("2测试用的文本2",2);
        XWPFParagraph paragraph1 = document.createParagraph();
        paragraph1.createRun().setText("普通的文本内容");
        //---------------------------------------------------------------------------------------//
        createDefaultHeader(document,"创建一个页眉");
//        createHeader(document,"公司全名","F:/poi-word-test/test.png");

        //---------------------------------------------------------------------------------------//
        XWPFParagraph paragraph2 = document.createParagraph();
        paragraph2.setIndentationFirstLine(2);
        paragraph2.setSpacingBefore(2);
        paragraph2.createRun().setText("道可道，非常道；名可名，非常名。无名天地之始，有名万物之母。故常无欲，以观其妙；常有欲，" +
                "以观其徼（jiào）。此两者同出而异名，同谓之玄，玄之又玄，众妙之门。" +
                "天下皆知美之为美，斯恶（è）已；皆知善之为善，斯不善已。故有无相生，难易相成，长短相较，高下相倾，音声相和（hè），前后相随。" +
                "是以圣人处无为之事，行不言之教，万物作焉而不辞，生而不有，为而不恃，功成而弗居。夫（fú）唯弗居，是以不去。" +
                "不尚贤，使民不争；不贵难得之货，使民不为盗；不见（xiàn）可欲，使民心不乱。是以圣人之治，虚其心，实其腹；弱其志，强其骨。" +
                "常使民无知无欲，使夫（fú）智者不敢为也。为无为，则无不治。" +
                "道冲而用之或不盈，渊兮似万物之宗。挫其锐，解其纷，和其光，同其尘。湛兮似或存，吾不知谁之子，象帝之先。");
        XWPFParagraph paragraph3 = document.createParagraph();
        paragraph3.setFirstLineIndent(2);
        paragraph3.createRun().setText("天地不仁，以万物为刍（chú）狗；圣人不仁，以百姓为刍狗。天地之间，其犹橐龠（tuóyuè）乎？虚而不屈，动而愈出。多言数（shuò）穷，不如守中。" +
                "谷神不死，是谓玄牝（pìn），玄牝之门，是谓天地根。绵绵若存，用之不勤。" +
                "天长地久。天地所以能长且久者，以其不自生，故能长生。是以圣人后其身而身先，外其身而身存。非以其无私邪（yé）？故能成其私。 " +
                "上善若水。水善利万物而不争，处众人之所恶（wù），故几（jī）于道。居善地，心善渊，与善仁，言善信，正善治，事善能，动善时。夫唯不争，故无尤。");
        //--------------------------------------------------------------------------------------//
        String text = "A B C D E F G H I J K L M N O P Q R S T U V W X Y Z 0 1 2 3 4 5 6 7 8 9 + - * / ~ ! @ # $ % ^ & * ( ) _ ? < > { } [ ]";
        document.createParagraph().createRun().setText("特殊符号-1:");
        XWPFRun xwpfRun = document.createParagraph().createRun();
        xwpfRun.setText(text);
        xwpfRun.setFontFamily("Wingdings 2");
        xwpfRun.setFontSize(16);
        xwpfRun.addCarriageReturn();
        document.createParagraph().createRun().setText("特殊符号-2:");
        XWPFRun xwpfRun1 = document.createParagraph().createRun();
        xwpfRun1.setText(text);
        xwpfRun1.setFontFamily("Wingdings 3");
        xwpfRun1.setFontSize(16);
        xwpfRun1.addCarriageReturn();
        document.createParagraph().createRun().setText("特殊符号-3:");
        XWPFRun xwpfRun2 = document.createParagraph().createRun();
        xwpfRun2.setText(text);
        xwpfRun2.setFontFamily("Wingdings");
        xwpfRun2.setFontSize(16);
        xwpfRun2.addCarriageReturn();
        document.createParagraph().createRun().setText("特殊符号-4:");
        XWPFRun xwpfRun3 = document.createParagraph().createRun();
        xwpfRun3.setText(text);
        xwpfRun3.setFontFamily("Symbol");
        xwpfRun3.setFontSize(16);
        xwpfRun3.addCarriageReturn();
        document.createParagraph().createRun().setText("特殊符号-5:");
        XWPFRun xwpfRun4 = document.createParagraph().createRun();
        xwpfRun4.setText(text);
        xwpfRun4.setFontFamily("Webdings");
        xwpfRun4.setFontSize(16);
        xwpfRun4.addCarriageReturn();
        FileOutputStream outputStream = new FileOutputStream("F:/poi-word-test/CreateTest.doc");
        document.write(outputStream);
        outputStream.close();
    }
    /**
     * 创建默认页眉
     *
     * @param docx XWPFDocument文档对象
     * @param text 页眉文本
     * @return 返回文档帮助类对象，可用于方法链调用
     */
    public static void createDefaultHeader(final XWPFDocument docx, final String text) throws IOException, InvalidFormatException {
        CTP ctp = CTP.Factory.newInstance();
        XWPFParagraph paragraph = new XWPFParagraph(ctp, docx);
        ctp.addNewR().addNewT().setStringValue(text);
        ctp.addNewR().addNewT().setSpace(SpaceAttribute.Space.PRESERVE);
        XWPFRun xwpfRun = paragraph.createRun();
        paragraph.addRun(xwpfRun);
        CTSectPr sectPr = docx.getDocument().getBody().isSetSectPr() ? docx.getDocument().getBody().getSectPr() : docx.getDocument().getBody().addNewSectPr();
        XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(docx, sectPr);
        XWPFHeader header = policy.createHeader(STHdrFtr.DEFAULT, new XWPFParagraph[] { paragraph });
        XWPFRun run = header.getParagraphs().get(0).getRuns().get(0);
        InputStream is = new FileInputStream("F:/poi-word-test/test.png");
        insertPicture(xwpfRun,docx,is,XWPFDocument.PICTURE_TYPE_PNG,20,30);
//        String data = header.addPictureData(is, XWPFDocument.PICTURE_TYPE_PNG);
//        ctp.addNewR().addNewT().setStringValue(data);

        header.setXWPFDocument(docx);
    }

    public static void createHeader(XWPFDocument doc, String orgFullName, String logoFilePath) throws Exception {
        /*
        * 对页眉段落作处理，使公司logo图片在页眉左边，公司全称在页眉右边
        * */
        CTSectPr sectPr = doc.getDocument().getBody().addNewSectPr();
        XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(doc,     sectPr);

        CTP ctp = CTP.Factory.newInstance();
        XWPFParagraph paragraph = new XWPFParagraph(ctp, doc);
        ctp.addNewR().addNewT().setStringValue(orgFullName);
        ctp.addNewR().addNewT().setSpace(SpaceAttribute.Space.PRESERVE);

        XWPFHeader header = headerFooterPolicy.createHeader(XWPFHeaderFooterPolicy.DEFAULT,new XWPFParagraph[]{paragraph});

        paragraph.setAlignment(ParagraphAlignment.LEFT);
        paragraph.setBorderBottom(Borders.THICK);

        CTTabStop tabStop = paragraph.getCTP().getPPr().addNewTabs().addNewTab();
        tabStop.setVal(STTabJc.RIGHT);
        int twipsPerInch =  1440;
        tabStop.setPos(BigInteger.valueOf(6 * twipsPerInch));

        XWPFRun run = paragraph.createRun();
//        setXWPFRunStyle(run,"新宋体",10);

        /*
        * 根据公司logo在ftp上的路径获取到公司到图片字节流
        * 添加公司logo到页眉，logo在左边
        * */
        if (logoFilePath!=null) {
//            String imgFile = FileUploadUtil.getLogoFilePath(logoFilePath);
//            byte[] bs = FtpUtil.downloadFileToIo(imgFile);
//            InputStream is = new ByteArrayInputStream(bs);
            InputStream is = new FileInputStream(logoFilePath);

            XWPFPicture picture = run.addPicture(is, XWPFDocument.PICTURE_TYPE_JPEG, "fileName", Units.toEMU(80), Units.toEMU(45));

            String blipID = "";
            for(XWPFPictureData picturedata : header.getAllPackagePictures()) { //这段必须有，不然打开的logo图片不显示
                blipID = header.getRelationId(picturedata);
            }
            picture.getCTPicture().getBlipFill().getBlip().setEmbed(blipID);
            run.addTab();
            is.close();
        }

        /*
         * 添加字体页眉，公司全称
         * 公司全称在右边
         * */
        if (orgFullName!=null) {
            run = paragraph.createRun();
            run.setText(orgFullName);
//            setXWPFRunStyle(run,"新宋体",10);
        }
    }


    //插入图片
    public static void insertPicture(XWPFRun run, XWPFDocument document, InputStream is, int type, int width, int height) {
        try {
            String picId = document.addPictureData(is, type);
            width = Units.toEMU(width);
            height = Units.toEMU(height);
            CTInline inline = run.getCTR().addNewDrawing().addNewInline();
            String picXml = "<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
                    "   <a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
                    "      <pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
                    "         <pic:nvPicPr>" +
                    "            <pic:cNvPr id=\"" + type + "\" name=\"Generated\"/>" +
                    "            <pic:cNvPicPr/>" +
                    "         </pic:nvPicPr>" +
                    "         <pic:blipFill>" +
                    "            <a:blip r:embed=\"" + picId + "\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>" +
                    "            <a:stretch>" +
                    "               <a:fillRect/>" +
                    "            </a:stretch>" +
                    "         </pic:blipFill>" +
                    "         <pic:spPr>" +
                    "            <a:xfrm>" +
                    "               <a:off x=\"0\" y=\"0\"/>" +
                    "               <a:ext cx=\"" + width + "\" cy=\"" + height + "\"/>" +
                    "            </a:xfrm>" +
                    "            <a:prstGeom prst=\"rect\">" +
                    "               <a:avLst/>" +
                    "            </a:prstGeom>" +
                    "         </pic:spPr>" +
                    "      </pic:pic>" +
                    "   </a:graphicData>" +
                    "</a:graphic>";
            XmlToken xmlToken = XmlToken.Factory.parse(picXml);
            inline.set(xmlToken);
            inline.setDistT(0);
            inline.setDistB(0);
            inline.setDistL(0);
            inline.setDistR(0);
            CTPositiveSize2D extent = inline.addNewExtent();
            extent.setCx(width);
            extent.setCy(height);
            CTNonVisualDrawingProps docPr = inline.addNewDocPr();
            docPr.setId(type);
            docPr.setName("Picture " + type);
            docPr.setDescr("Generated");
        } catch (Exception xe) {
            xe.printStackTrace();
        }
    }
}
