package priv.viking.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.Calendar;
import java.util.Date;

/**
 * Created By Viking on 2021/9/8
 * poi excel操作官方文档链接:https://poi.apache.org/components/spreadsheet/quick-guide.html#NewWorkbook
 */
public class ExportExcel {
    public static void create() {
        // 创建新的工作薄
    //    Workbook wb = new HSSFWorkbook();
        Workbook wb = new XSSFWorkbook();

        // 通过这个工具方法可以帮我们创建一个安全可用的文档名称
        String safeName = WorkbookUtil.createSafeSheetName("[O'Brien's sales*?]"); // returns " O'Brien's sales   "

        // 创建文档
        // 文档名称不能超过31个字符，且不能包含以下任一字符：0x0000 0x0003 : \ * ? / [ ]
        Sheet sheet = wb.createSheet();

        CreationHelper createHelper = wb.getCreationHelper();
        // 创建一行，并在其中放置一些单元格，行的下标是基于0开始的
        Row row = sheet.createRow(0);
        // 创建一个单元格并设置一个值
        Cell cell = row.createCell(0);
        cell.setCellValue(1);
        // 写一行
        row.createCell(1).setCellValue(1.2);
        row.createCell(2).setCellValue(
                createHelper.createRichTextString("This is a string"));
        row.createCell(3).setCellValue(true);

        //从工作簿中创建一个新的单元格样式，否则最终可能会修改内置样式，不仅影响这个单元格，还会影响其他单元格。
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setDataFormat(
                createHelper.createDataFormat().getFormat("m/d/yy h:mm"));
        cell = row.createCell(1);
        cell.setCellValue(new Date());
        cell.setCellStyle(cellStyle);

        // 处理不同类型的单元格
        Row row1 = sheet.createRow(1);
        row1.createCell(0).setCellValue(1.1);
        row1.createCell(1).setCellValue(new Date());
        row1.createCell(2).setCellValue(Calendar.getInstance());
        row1.createCell(3).setCellValue("a string");
        row1.createCell(4).setCellValue(true);
        row1.createCell(5).setCellType(CellType.ERROR);

        ////////////////////////////////////////////////////////////////
        // 当打开工作簿时，无论是.xls HSSFWorkbook还是.xlsx XSSFWorkbook，工作簿都可以从File或InputStream加载。
        // 使用File对象可以降低内存消耗，而InputStream需要更多内存，因为它必须缓冲整个文件。
        // 如果使用WorkbookFactory，那么两者都很容易使用
        try {
            // Use a file
            Workbook wb1 = WorkbookFactory.create(new File("MyExcel.xls"));
            // Use an InputStream, needs more memory
            Workbook wb2 = WorkbookFactory.create(new FileInputStream("MyExcel.xlsx"));
        } catch (InvalidFormatException | IOException e) {
            e.printStackTrace();
        }

        /////////////////////////////////////////////////////////////
        try {
            // 如果直接使用HSSFWorkbook或XSSFWorkbook，通常应该通过POIFSFileSystem或OPCPackage，以完全控制生命周期(包括关闭文件时完成):
            // HSSFWorkbook, File
            POIFSFileSystem fs = new POIFSFileSystem(new File("file.xls"));
            HSSFWorkbook wb2 = new HSSFWorkbook(fs.getRoot(), true);

            fs.close();
            // HSSFWorkbook, InputStream, needs more memory
            POIFSFileSystem fs1 = new POIFSFileSystem(new FileInputStream("file.xls"));
            HSSFWorkbook wb3 = new HSSFWorkbook(fs.getRoot(), true);
            // XSSFWorkbook, File
            OPCPackage pkg = OPCPackage.open(new File("file.xlsx"));
            XSSFWorkbook wb4 = new XSSFWorkbook(pkg);

            pkg.close();
            // XSSFWorkbook, InputStream, needs more memory
            OPCPackage pkg1 = OPCPackage.open(new FileInputStream("Mfile.xls"));
            XSSFWorkbook wb5 = new XSSFWorkbook(pkg);

            pkg.close();
        } catch (InvalidFormatException | IOException e) {
            e.printStackTrace();
        }
        //////////////////////////////////////////////////////////////////////////////
        try {

            // 图片处理
            // 图像是绘图支持的一部分。要添加图像，只需在绘图父结点上调用createPicture()。在撰写本文时，支持以下类型:PNG、JPG、DIB
            // 需要注意的是，一旦将图像添加到工作表中，任何现有的绘图都可能被删除。
            // 将图片数据添加到此工作簿
            InputStream is = new FileInputStream("image1.jpeg");
            byte[] bytes = IOUtils.toByteArray(is);
            int pictureIdx = wb.addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);
            is.close();
            CreationHelper helper = wb.getCreationHelper();
             //create sheet
//            Sheet sheet = wb.createSheet();
            // Create the drawing patriarch.  This is the top level container for all shapes.
            // 创建绘图管理器。这是所有形状的顶级容器。
            @SuppressWarnings("rawtypes")
            Drawing drawing = sheet.createDrawingPatriarch();
             // 添加图片形状
            ClientAnchor anchor = helper.createClientAnchor();
            //set top-left corner of the picture,subsequent call of Picture#resize() will operate relative to it
            // 设置图片的左上角定位点，随后调用的图片#resize()将相对于它进行操作
            // Picture.resize()仅适用于JPEG和PNG。其他格式还不支持。
            anchor.setCol1(3);
            anchor.setRow1(2);
            Picture pict = drawing.createPicture(anchor, pictureIdx);
            //auto-size picture relative to its top-left corner
            // 自动大小的图片相对于它的左上角定位点
            pict.resize();
            //save workbook
            String file = "picture.xls";
            if (wb instanceof XSSFWorkbook) file += "x";
        } catch (IOException e) {
            e.printStackTrace();
        }

        try (
            // 导出到文件
            OutputStream fileOut = new FileOutputStream("workbook.xlsx")) {
            wb.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) {
        String safeName = WorkbookUtil.createSafeSheetName("[O'Brien's sales*?]AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV");
        System.out.println(safeName);
    }

    public static String getPath(String an){
        String rootPath="https://zqz-shangbiao.obs.cn-east-3.myhuaweicloud.com";
        String path;
        try {
            path=rootPath+"/"+an.substring(0,1)+"/"+getPathSize(an)+"/"+an+".png";
        }catch (Exception e){
            String size = an.replaceAll("[A-Za-z]","");
            path=rootPath+"/"+"_"+an.substring(an.length()-1)+"/"+getPathSize(size)+"/"+an+".png";
        }
        return path;
    }
    private static String getPathSize(String an){
        if(an == null || an.equals("")) return "";
        return ((Integer.parseInt(an)/5000)+1)+"";
    }

    /**
     * 根据地址获得数据的字节流
     * @param strUrl 网络连接地址
     * @return
     */
    public static byte[] getImageFromNetByUrl(String strUrl){
        System.out.println(strUrl);
        try {
            URL url = new URL(strUrl);
            HttpURLConnection conn = (HttpURLConnection)url.openConnection();
            conn.setRequestMethod("GET");
            conn.setConnectTimeout(5 * 1000);
            InputStream inStream = conn.getInputStream();//通过输入流获取图片数据
            byte[] btImg = readInputStream(inStream);//得到图片的二进制数据
            return btImg;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }
    /**
     * 从输入流中获取数据
     * @param inStream 输入流
     * @return
     * @throws Exception
     */
    public static byte[] readInputStream(InputStream inStream) throws Exception{
        ByteArrayOutputStream outStream = new ByteArrayOutputStream();
        byte[] buffer = new byte[1024];
        int len = 0;
        while( (len=inStream.read(buffer)) != -1 ){
            outStream.write(buffer, 0, len);
        }
        inStream.close();
        return outStream.toByteArray();
    }

}
