/**
 * Copyright (C), 2015-2018, 华电408有限公司
 * FileName: WriteExcel
 * Author:   宣佚
 * Date:     2018/8/1 0001 下午 20:48
 * Description: 对excel进行写入
 * History:
 * <author>          <time>          <version>          <desc>
 * 作者姓名           修改时间           版本号              描述
 */
package com.iflytek.excel;

/**
 * 〈一句话功能简述〉<br> 
 * 〈对excel进行写入〉
 *
 * @author 宣佚
 * @create 2018/8/1 0001
 * @since 1.0.0
 */

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Vector;

public class WriteExcel {

    private String pathname;
    private Workbook workbook;
    private Sheet sheet1;

    /**使用栗子
     * WriteExcel excel = new WriteExcel("F:\\javaa\\MscDemo\\excel\\myexcel.xlsx");
     * excel.write(new String[]{"1","2"}, 0);//在第1行第1个单元格写入1,第一行第二个单元格写入2
     */
    public void write(Vector writeStrings, int rowNumber) throws Exception {
        //将内容写入指定的行号中
        Row row = sheet1.createRow(rowNumber);
        //遍历整行中的列序号
        for (int j = 0; j < writeStrings.size(); j++) {
            //根据行指定列坐标j,然后在单元格中写入数据
            Cell cell = row.createCell(j);
            cell.setCellValue(writeStrings.get(j).toString());
        }
        OutputStream stream = new FileOutputStream(pathname);
        workbook.write(stream);
        stream.close();
    }
    /**
     * 功能描述: <br>
     * 〈构造函数，初始化workbook和worksheet，根据创建的excel类型初始化workbook〉
     *
     * @param excelPath
     * @return:
     * @since: 1.0.0
     * @Author:宣佚
     * @Date: 2018/8/1 0001 下午 20:53
     */
    public WriteExcel(String excelPath) throws Exception {
        //在excelPath中需要指定具体的文件名(需要带上.xls或.xlsx的后缀)
        this.pathname = excelPath;
        String fileType = excelPath.substring(excelPath.lastIndexOf(".") + 1, excelPath.length());
        if(isExist(pathname)){
            if(fileType.equals("xls")){
                //如果是.xls,就new HSSFWorkbook(file)
                workbook=new HSSFWorkbook(new FileInputStream(pathname));
            }else if(fileType.equals("xlsx")){
                //如果是.xlsx,就new XSSFWorkbook(file)
                workbook=new XSSFWorkbook(new FileInputStream(pathname));
            }else{
                throw new Exception("文档格式后缀不正确!!！");
            }
            sheet1=workbook.getSheet("sheet1");
        }else {
            //创建文档对象
            if (fileType.equals("xls")) {
                //如果是.xls,就new HSSFWorkbook()
                workbook = new HSSFWorkbook();
            } else if (fileType.equals("xlsx")) {
                //如果是.xlsx,就new XSSFWorkbook()
                workbook = new XSSFWorkbook();
            } else {
                throw new Exception("文档格式后缀不正确!!！");
            }
            // 创建表sheet
            sheet1 = workbook.createSheet("sheet1");
        }
    }

    /**
     * 功能描述: <br>
     * 〈判断excel文件是否存在〉
     *
     * @param excelPath
     * @return:boolean
     * @since: 1.0.0
     * @Author:宣佚
     * @Date: 2018/8/1 0001 下午 21:19
     */
    public boolean isExist(String excelPath){
        File file=new File(excelPath);
        return file.exists();
    }

    public static void main(String args[]){
        try {
            WriteExcel writeExcel=new WriteExcel("F:\\javaa\\MscDemo\\excel\\myExcel1.xlsx");
            Vector strArray=new Vector();
            strArray.add("李宣佚");
            strArray.add("李建国");

            writeExcel.write(strArray,1);
        }catch (Exception e){
            e.printStackTrace();
        }
    }
}

