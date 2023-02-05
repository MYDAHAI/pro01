package com.yr.springboot_poi.poi;

import com.sun.media.sound.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.beans.factory.annotation.Autowired;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

public class ReadExcel {

    public static void main(String[] args) {
        List<User> list = new ArrayList<>();
        String path = "C:\\Users\\hai阳\\Desktop\\a.xlsx";
        new ReadExcel().readExcelToObj(path,list);
        System.out.println(list);
    }

    /**
     * 读取excel数据
     *
     * @param path
     */
    private void readExcelToObj(String path,List<User> list) {
        Workbook wb = null;
        try {
            wb = WorkbookFactory.create(new File(path));  //得到Excel文件对象
            int a = wb.getNumberOfSheets();  //得到表格数
            for(int i = 0;i<a;i++)
            {
                readExcel(wb, i, 0, 0,list);
            }
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 读取excel文件
     *
     * @param wb
     * @param sheetIndex
     *            sheet页下标：从0开始
     * @param startReadLine
     *            开始读取的行:从0开始
     * @param tailLine
     *            去除最后读取的行
     */
    @SuppressWarnings("deprecation") //允许您选择性地取消特定代码段（即，类或方法）中的警告
    private void readExcel(Workbook wb, int sheetIndex, int startReadLine, int tailLine,List<User> list) {

        Sheet sheet = wb.getSheetAt(sheetIndex); //得到第sheetIndex个表格的表格对象
        Row row = null;

        for (int i = startReadLine; i < sheet.getLastRowNum() - tailLine + 1; i++) {
            row = sheet.getRow(i);  //Row 行对象
            if (null != row) {
                User user = new User();
                for (Cell c : row) {  //Cell 行中的格子
					/*if(c.getCellTypeEnum() == CellType.STRING)
					{
						System.out.print(c.getStringCellValue() + "　　　　 ");
					}
					else if(c.getCellTypeEnum() == CellType.NUMERIC)
					{
						System.out.print(c.getNumericCellValue() + "　　　　 ");
					}*/

                    //需要先判断是否有合并      getColumnIndex 获取该列的索引下标
                    boolean isMerge = isMergedRegion(sheet, i, c.getColumnIndex());
                    // 判断是否具有合并单元格
                    if (null != c) {
                        String excelName = ""; //得到列的属性名
                        if(i != 0){
                            Row row1 = sheet.getRow(0); //得到表中第一行的行对象
                            int columnIndex = c.getColumnIndex(); //该列的索引
                            if(row1 != null){
                                for(Cell ce : row1){
                                    if(ce.getColumnIndex() == columnIndex){
                                        excelName = ce.getStringCellValue();
                                    }
                                }
                            }
                        }

                        if (isMerge) {
                            //需要通过自定义方法来获取合并单元格的值  getRowNum 得到当前行的索引
                            Cell cel = getCellObject(sheet, row.getRowNum(), c.getColumnIndex());
                            if(cel.getCellType() == CellType.STRING){ //判断类型是不是字符串类型
                                System.out.print(cel.getStringCellValue() + "　　　　 ");
                                if("name".equals(excelName)) user.setName(cel.getStringCellValue());
                            }else if(cel.getCellType() == CellType.NUMERIC){ //判断类型是不是数字类型
                                System.out.print(cel.getNumericCellValue() + "　　　　 ");
                                if("id".equals(excelName)) user.setId((int) cel.getNumericCellValue());
                                if("age".equals(excelName)) user.setAge((int) cel.getNumericCellValue());
                            }
                        } else {
                            if(c.getCellType() == CellType.STRING){ //判断类型是不是字符串类型
                                System.out.print(c.getStringCellValue() + "　　　　 ");
                                if("name".equals(excelName)) user.setName(c.getStringCellValue());
                            }else if(c.getCellType() == CellType.NUMERIC){ //判断类型是不是数字类型
                                System.out.print(c.getNumericCellValue() + "　　　　 ");
                                if("id".equals(excelName)) user.setId((int) c.getNumericCellValue());
                                if("age".equals(excelName)) user.setAge((int) c.getNumericCellValue());
                            }
                        }
                    }
                }
                System.out.println();
                if(i != 0)list.add(user);
            }
        }
    }


    /**
     * 获得合并单元格的列对象
     * @param sheet
     * @param row
     *            行下标
     * @param column
     *            列下标
     * @return
     */
    public Cell getCellObject(Sheet sheet, int row, int column) {
        //getNumMergedRegions 获得该sheet所有合并单元格数量
        int sheetMergeCount = sheet.getNumMergedRegions();

        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress ca = sheet.getMergedRegion(i);// 获得合并区域
            int firstColumn = ca.getFirstColumn();// 合并单元格CELL起始列
            int lastColumn = ca.getLastColumn();// 合并单元格CELL结束列
            int firstRow = ca.getFirstRow();// 合并单元格CELL起始行
            int lastRow = ca.getLastRow();// 合并单元格CELL结束行

            if (row >= firstRow && row <= lastRow) {// 判断该单元格是否是在合并单元格中

                if (column >= firstColumn && column <= lastColumn) {
                    Row fRow = sheet.getRow(firstRow);
                    Cell fCell = fRow.getCell(firstColumn);
                    return fCell;
                }
            }
        }

        return null;
    }


    /**
     * 获取合并单元格的值
     *
     * @param sheet
     * @param row
     *            行下标
     * @param column
     *            列下标
     * @return
     */
    public String getMergedRegionValue(Sheet sheet, int row, int column) {
        //getNumMergedRegions 获得该sheet所有合并单元格数量
        int sheetMergeCount = sheet.getNumMergedRegions();

        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress ca = sheet.getMergedRegion(i);// 获得合并区域
            int firstColumn = ca.getFirstColumn();// 合并单元格CELL起始列
            int lastColumn = ca.getLastColumn();// 合并单元格CELL结束列
            int firstRow = ca.getFirstRow();// 合并单元格CELL起始行
            int lastRow = ca.getLastRow();// 合并单元格CELL结束行

            if (row >= firstRow && row <= lastRow) {// 判断该单元格是否是在合并单元格中

                if (column >= firstColumn && column <= lastColumn) {
                    Row fRow = sheet.getRow(firstRow);
                    Cell fCell = fRow.getCell(firstColumn);
                    return getCellValue(fCell);
                }
            }
        }

        return null;
    }


    /**
     * 判断合并了行
     *
     * @param sheet
     * @param row
     * @param column
     * @return
     */
    @SuppressWarnings("unused")
    private boolean isMergedRow(Sheet sheet, int row, int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if (row == firstRow && row == lastRow) {
                if (column >= firstColumn && column <= lastColumn) {
                    return true;
                }
            }
        }
        return false;
    }

    /**
     * 判断指定的单元格是否是合并单元格
     *
     * @param sheet
     * @param row
     *            行下标
     * @param column
     *            列下标
     * @return
     */
    private boolean isMergedRegion(Sheet sheet, int row, int column) {
        //getNumMergedRegions 获得该sheet所有合并单元格数量
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i); // 获得合并区域
            int firstColumn = range.getFirstColumn();// 合并单元格CELL起始列
            int lastColumn = range.getLastColumn();// 合并单元格CELL结束列
            int firstRow = range.getFirstRow();// 合并单元格CELL起始行
            int lastRow = range.getLastRow();// 合并单元格CELL结束行
            if (row >= firstRow && row <= lastRow) {// 判断该单元格是否是在合并单元格中
                if (column >= firstColumn && column <= lastColumn) {
                    return true;
                }
            }
        }
        return false;
    }

    /**
     * 判断sheet页中是否含有合并单元格
     *
     * @param sheet
     * @return
     */
    @SuppressWarnings("unused")
    private boolean hasMerged(Sheet sheet) {
        return sheet.getNumMergedRegions() > 0 ? true : false;
    }

    /**
     * 合并单元格
     *
     * @param sheet
     * @param firstRow
     *            开始行
     * @param lastRow
     *            结束行
     * @param firstCol
     *            开始列
     * @param lastCol
     *            结束列
     */
    @SuppressWarnings("unused")
    private void mergeRegion(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
        sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
    }

    /**
     * 获取单元格的值
     *
     * @param cell
     * @return
     */
    @SuppressWarnings("deprecation")
    public String getCellValue(Cell cell) {

        if (cell == null)
            return "";

        if (cell.getCellType() == CellType.STRING) {

            return cell.getStringCellValue();

        } else if (cell.getCellType() == CellType.BOOLEAN) {

            return String.valueOf(cell.getBooleanCellValue());

        } else if (cell.getCellType() == CellType.FORMULA) {

            return cell.getCellFormula();

        } else if (cell.getCellType() == CellType.NUMERIC) {

            return String.valueOf(cell.getNumericCellValue());

        }
        return "";
    }

}
