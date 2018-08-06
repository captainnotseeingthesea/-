/**
 * Copyright (C), 2015-2018, 华电408有限公司
 * FileName: ReadExcel
 * Author:   宣佚
 * Date:     2018/8/1 0001 下午 19:17
 * Description: 对Excel进行读取
 * History:
 * <author>          <time>          <version>          <desc>
 * 作者姓名           修改时间           版本号              描述
 */
package com.iflytek.excel;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Vector;

/**
 * 〈一句话功能简述〉<br> 
 * 〈对Excel进行读取〉
 *
 * @author 宣佚
 * @create 2018/8/1 0001
 * @since 1.0.0
 */
public class ReadExcel {
    private HSSFSheet hssfSheet;//.xls
    private XSSFSheet xssfSheet;//.xlsx

    public int getAllRowNumber() {
        return xssfSheet.getLastRowNum();
    }

    /*读取 excel 下标为 rowNumber 的那一行的全部数据*/
    public Vector readLine(int rowNumber) {
        XSSFRow row = xssfSheet.getRow(rowNumber);
        if (row != null) {
            Vector resultStr=new Vector();
            for (int i = 0; i < row.getLastCellNum(); i++) {
                String value="";//获取的单元格的内容
                Cell cell=row.getCell(i);//获取单元格对象
                if (cell!=null){//判断单元格是否为空
                    switch (cell.getCellType()) {
                        case HSSFCell.CELL_TYPE_NUMERIC: // 数字
                            //如果为时间格式的内容
                            if (HSSFDateUtil.isCellDateFormatted(cell)) {
                                //注：format格式 yyyy-MM-dd hh:mm:ss 中小时为12小时制，若要24小时制，则把小h变为H即可，yyyy-MM-dd HH:mm:ss
                                SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd hh:mm:ss");
                                value=sdf.format(HSSFDateUtil.getJavaDate(cell.
                                        getNumericCellValue()));
                                break;
                            } else {
                                value = new DecimalFormat("0").format(cell.getNumericCellValue());
                            }
                            break;
                        case HSSFCell.CELL_TYPE_STRING: // 字符串
                            value = cell.getStringCellValue();
                            break;
                        case HSSFCell.CELL_TYPE_BOOLEAN: // Boolean
                            value = cell.getBooleanCellValue() + "";
                            break;
                        case HSSFCell.CELL_TYPE_FORMULA: // 公式
                            value = cell.getCellFormula() + "";
                            break;
                        case HSSFCell.CELL_TYPE_BLANK: // 空值
                            value = "";
                            break;
                        case HSSFCell.CELL_TYPE_ERROR: // 故障
                            value = "非法字符";
                            break;
                        default:
                            value = "未知类型";
                            break;
                    }
                }

                resultStr.add(value);
            }
            return resultStr;
        }
        return null;
    }


    public ReadExcel(String excelPath) throws Exception {
        String fileType = excelPath.substring(excelPath.lastIndexOf(".") + 1, excelPath.length());
        // 创建工作文档对象
        InputStream in = new FileInputStream(excelPath);
        HSSFWorkbook hssfWorkbook = null;//.xls
        XSSFWorkbook xssfWorkbook = null;//.xlsx
        //根据后缀创建读取不同类型的excel
        if (fileType.equals("xls")) {
            hssfWorkbook = new HSSFWorkbook(in);//它是专门读取.xls的
        } else if (fileType.equals("xlsx")) {
            xssfWorkbook = new XSSFWorkbook(in);//它是专门读取.xlsx的
        } else {
            throw new Exception("文档格式后缀不正确!!！");
        }
        /*这里默认只读取第 1 个sheet*/
        if (hssfWorkbook != null) {
            this.hssfSheet = hssfWorkbook.getSheetAt(0);
        } else if (xssfWorkbook != null) {
            this.xssfSheet = xssfWorkbook.getSheetAt(0);
        }
    }

    public static void main(String args[]){
        try {
            ReadExcel readExcel=new ReadExcel("F:\\javaa\\MscDemo\\excel\\1.xlsx");
            Vector resultStr=readExcel.readLine(0);
            for(int i=0;i<resultStr.size();i++){
                System.out.println(resultStr.get(i));
            }
            System.out.println(readExcel.getAllRowNumber());
        }catch (Exception e){
            e.printStackTrace();
        }
    }
}
