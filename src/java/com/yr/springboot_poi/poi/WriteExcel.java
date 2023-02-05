package com.yr.springboot_poi.poi;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;

public class WriteExcel {

    /**
     * @param args
     */
    @SuppressWarnings("deprecation")
    public static void main(String[] args) throws Exception {
        // 创建Excel的工作书册 Workbook,对应到一个excel文档
        XSSFWorkbook wb = new XSSFWorkbook();

        // 创建Excel的工作sheet,对应到一个excel文档的tab

        XSSFSheet sheet = wb.createSheet("sheet1"); //创建一个表格
//        wb.createSheet("唐骞");

        // 设置excel每列宽度
        sheet.setColumnWidth(0, 1000); //第一个格子宽度
        sheet.setColumnWidth(1, 8500);//第二个格子宽度

        // 创建字体样式
        // XSSFFont font = wb.createFont();
                /*font.setFontName("Verdana");
                font.setBoldweight((short) 100);
                font.setFontHeight((short) 300);
                font.setColor(HSSFColor.BLUE.index);*/

        //ROW  CELL
        // 创建单元格样式
        //XSSFCellStyle style = wb.createCellStyle();
               /* style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
                style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
                style.setFillForegroundColor(XSSFColor.LIGHT_TURQUOISE.index);
                style.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);*/

        // 设置边框
               /* style.setBottomBorderColor(HSSFColor.RED.index);
                style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
                style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
                style.setBorderRight(HSSFCellStyle.BORDER_THIN);
                style.setBorderTop(HSSFCellStyle.BORDER_THIN);*/

        //style.setFont(font);// 设置字体

        // 创建Excel的sheet的一行
        XSSFRow row = sheet.createRow(0);

        // row.setHeight((short) 500);// 设定行的高度
        // 创建一个Excel的单元格
        XSSFCell cell0 = row.createCell(0);
        XSSFCell cell1 = row.createCell(1);

        // 合并单元格(startRow，endRow，startColumn，endColumn)   起始行号，终止行号， 起始列号，终止列号
        sheet.addMergedRegion(new CellRangeAddress(0, 2, 0, 2));


        //springmvc 导入，导出

        // 给Excel的单元格设置样式和赋值
        // cell.setCellStyle(style);
        cell0.setCellValue("hello world");
        cell1.setCellValue("hello worldxxxx");

        // 设置单元格内容格式
        // XSSFCellStyle style1 = wb.createCellStyle();
        // style1.setDataFormat(HSSFDataFormat.getBuiltinFormat("h:mm:ss"));

        // style1.setWrapText(true);// 自动换行

        //row = sheet.createRow(1);

        // 设置单元格的样式格式

        //cell = row.createCell(0);
        // cell.setCellStyle(style1);
        // cell.setCellValue(new Date());

        // 创建超链接
            /*    HSSFHyperlink link = new HSSFHyperlink(HSSFHyperlink.LINK_URL);
                link.setAddress("http://www.baidu.com");
                cell = row.createCell(1);
                cell.setCellValue("百度");
                cell.setHyperlink(link);// 设定单元格的链接
             */
        FileOutputStream os = new FileOutputStream("C:\\Users\\hai阳\\Desktop\\b.xlsx");
        wb.write(os);
        os.close();

    }

}
