//package priv.viking.excel;
//
//import java.io.File;
//import java.io.FileOutputStream;
//import java.net.URLEncoder;
//import java.util.List;
//import javax.servlet.http.HttpServletResponse;
//import org.apache.poi.hssf.usermodel.HSSFCellStyle;
//import org.apache.poi.hssf.usermodel.HSSFFont;
//import org.apache.poi.hssf.util.HSSFColor;
//import org.apache.poi.ss.usermodel.CellStyle;
//import org.apache.poi.ss.usermodel.FillPatternType;
//import org.apache.poi.ss.usermodel.Font;
//import org.apache.poi.ss.usermodel.IndexedColors;
//import org.apache.poi.xssf.streaming.SXSSFCell;
//import org.apache.poi.xssf.streaming.SXSSFRow;
//import org.apache.poi.xssf.streaming.SXSSFSheet;
//import org.apache.poi.xssf.streaming.SXSSFWorkbook;
//import org.apache.poi.xssf.usermodel.XSSFRichTextString;
///**
// * Created By Viking on 2021/9/8
// */
//public class ExcelExport {
//
//    private ExcelExport INSTANSE = new ExcelExport ();
//
//    /**
//     * description: 导出数据excel
//     * @param sheetName
//     * @param headers
//     * @param dataList
//     * @param destFile
//     * @return void
//     * @version v1.0
//     * @author w
//     * @date 2020年3月30日 下午2:23:39
//     */
//    public static void export(String sheetName, String[] headers, List<List<Object>> dataList, File destFile) throws Exception {
//        SXSSFWorkbook workbook = new SXSSFWorkbook();
//        createSheet(sheetName, headers, dataList, workbook);
//        workbook.write(new FileOutputStream(destFile));
//    }
//
//
//    /**
//     * description: 导出excel --- 支持web
//     * @param sheetName  sheet表名字
//     * @param headers 表头
//     * @param dataList 表数据
//     * @param fileName  导出文件名
//     * @param response
//     * @return void
//     * @version v1.0
//     * @author w
//     * @date 2020年3月31日 下午2:48:46
//     */
//    public static void export(String sheetName , String[] headers , List<List<Object>> dataList ,String fileName
//            , HttpServletResponse response) throws Exception {
//        SXSSFWorkbook workbook = new SXSSFWorkbook();
//        createSheet(sheetName, headers, dataList, workbook);
//        response.reset();
//        response.setContentType("application/octet-stream; charset=utf-8");
//        response.setHeader("Content-Disposition", "attachment; filename="+URLEncoder.encode(fileName ,"UTF-8"));
//        workbook.write(response.getOutputStream());
//        // 删除临时文件
//        workbook.dispose();
//    }
//
//    /**
//     * description: 创建sheet表格
//     * @param sheetName  表sheet 名字
//     * @param headers  表头
//     * @param dataList 表数据
//     * @param wb
//     * @return void
//     * @version v1.0
//     * @author w
//     * @date 2020年3月30日 下午2:33:39
//     */
//    public static void createSheet(String sheetName , String[] headers , List<List<Object>> dataList , SXSSFWorkbook wb) {
//        SXSSFSheet sheet = wb.createSheet(sheetName);
//        // 设置表头和单元格格式
//        CellStyle headStyle = setHeaderStyle(wb);
//        CellStyle bodyStyle = setBodyStyle(wb);
//        // 创建表头和单元格数据
//        createHeader(headers, sheet, headStyle);
//        createBody(dataList, sheet, bodyStyle);
//    }
//
//    /**
//     * description: 创建表头
//     * @param headers
//     * @param sheet
//     * @param headStyle
//     * @return void
//     * @version v1.0
//     * @author w
//     * @date 2020年3月30日 下午3:03
//     */
//    private static void createHeader(String[] headers, SXSSFSheet sheet, CellStyle headStyle) {
//        SXSSFRow row = sheet.createRow(0);
//        row.setHeightInPoints(16F);
//        for (int i = 0; i < headers.length; i++) {
//            // 创建单元格
//            SXSSFCell cell = row.createCell(i);
//            cell.setCellStyle(headStyle);
//            XSSFRichTextString text = new XSSFRichTextString(headers[i]);
//            cell.setCellValue(text);
//            sheet.trackAllColumnsForAutoSizing();
//            sheet.autoSizeColumn(i);
//        }
//    }
//
//    /**
//     * description: 表格中填充数据
//     * @param dataList
//     * @param sheet
//     * @param bodyStyle
//     * @return void
//     * @version v1.0
//     * @author w
//     * @date  2020年3月30日 下午3:13
//     */
//    private static void createBody(List<List<Object>> dataList, SXSSFSheet sheet, CellStyle bodyStyle) {
//        for (int i = 0; i < dataList.size(); i++) {
//            // 从第二行开始，第一行做表头
//            SXSSFRow row = sheet.createRow(i+1);
//            List<Object> rowList = dataList.get(i);
//            for (int j = 0; j < rowList.size(); j++) {
//                SXSSFCell cell = row.createCell(j);
//                cell.setCellStyle(bodyStyle);
//                XSSFRichTextString text = new XSSFRichTextString(rowList.get(j).toString());
//                cell.setCellValue(text);
//                sheet.trackAllColumnsForAutoSizing();
//                sheet.autoSizeColumn(i);
//            }
//        }
//    }
//
//    /**
//     * description: 设置单元格内容样式
//     * @param wb
//     * @return HSSFCellStyle
//     * @version v1.0
//     * @author w
//     * @date 2020年3月30日 下午2:42:39
//     */
//    private static CellStyle setBodyStyle(SXSSFWorkbook wb) {
//        // 设置表格单元格格式
//        CellStyle style = wb.createCellStyle();
//        style.setFillForegroundColor(HSSFColor.WHITE.index);
//        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
//        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
//        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
//        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
//        style.setBorderTop(HSSFCellStyle.BORDER_THIN);
//        style.setAlignment(HSSFCellStyle.ALIGN_LEFT);
//
//        // 设置字体格式
//        Font font = wb.createFont();
//        font.setFontName("微软雅黑");
//        font.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
//        style.setFont(font);
//        return style;
//    }
//
//    /**
//     * description: 设置表头样式
//     * @param wb
//     * @return
//     * @return HSSFCellStyle
//     * @version v1.0
//     * @author w
//     * @date 2020年3月30日 下午2:38:39
//     */
//    private static CellStyle setHeaderStyle(SXSSFWorkbook wb) {
//        // 设置表格单元格格式
//        CellStyle style = wb.createCellStyle();
//        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
//        style.setBorderRight(CellStyle.BORDER_THIN);
//        style.setRightBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
//        style.setBorderLeft(CellStyle.BORDER_THIN);
//        style.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
//        style.setBorderTop(CellStyle.BORDER_THIN);
//        style.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
//        style.setBorderBottom(CellStyle.BORDER_THIN);
//        style.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
//
//        // 设置字体格式
//        Font font = wb.createFont();
//        font.setFontName("微软雅黑");
//        font.setFontHeightInPoints((short) 10);
//        style.setFont(font);
//        return style;
//    }
//
//
//}
