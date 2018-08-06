/**
 * Copyright (C), 2015-2018, 华电408有限公司
 * FileName: ExcelTable
 * Author:   宣佚
 * Date:     2018/8/2 0002 上午 9:13
 * Description:
 * History:
 * <author>          <time>          <version>          <desc>
 * 作者姓名           修改时间           版本号              描述
 */
package com.iflytek.view;

import com.iflytek.excel.ReadExcel;
import com.iflytek.excel.WriteExcel;

import javax.swing.*;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.util.HashMap;
import java.util.Vector;

/**
 * 〈一句话功能简述〉<br> 
 * 〈〉
 *
 * @author 宣佚
 * @create 2018/8/2 0002
 * @since 1.0.0
 */
public class ExcelTable {

    private JPanel mainPanel;
    private JTable excel;
    private JScrollPane tableScrolPane;
    private JButton integrate;
    private JButton fileChooser;
    private JLabel fileName;
    private JButton arrange;//保存文件的按钮
    private JButton saveFile;
    private JLabel referFile;
    private Vector data;//excel显示的数据
    private Vector title;//excel的表头
    private HashMap map;//条目对应的Hash值，vector对应的位置

    public ExcelTable(){
        DefaultTableCellRenderer rTable=new DefaultTableCellRenderer();//创建渲染表格内容的渲染器
        rTable.setHorizontalAlignment(JLabel.CENTER);
        excel.setDefaultRenderer(Object.class,rTable);//设置表格内部的内容居中
        excel.getTableHeader().setFont(new Font("叶根友毛笔行书2.0",Font.BOLD,18));

        integrate.addActionListener(new ActionListener() {//汇总文件
            @Override
            public void actionPerformed(ActionEvent e) {
                JFileChooser jfc=new JFileChooser(new File("."));
                jfc.setMultiSelectionEnabled(true);
                jfc.showDialog(new JLabel(), "选择");
                File file[]=jfc.getSelectedFiles();
                HashMap hashMap=new HashMap();//汇总表时的映射
                int hashIndex=0;//hash对应vector的映射
                Vector integrateData=new Vector();
                if(file!=null){
                    for(int i=0;i<file.length;i++){
                        String filePath=file[i].getAbsolutePath();
                        String fileType = filePath.substring(filePath.lastIndexOf(".") + 1, filePath.length());
                        if(fileType.equals("xls")||fileType.equals("xlsx")){
                            try {
                                ReadExcel readExcel=new ReadExcel(filePath);
                                for (int j=1;j<=readExcel.getAllRowNumber();j++){
                                    Vector row=readExcel.readLine(j);
                                    String str=row.get(1).toString()+row.get(2);//条目名称与规格合并
                                    if(hashMap.get(str)==null){
                                        row.set(0,hashIndex+1);
                                        integrateData.add(row);
                                        hashMap.put(str,hashIndex++);
                                    }else {
                                        int index=Integer.parseInt(hashMap.get(str).toString());
                                        ((Vector)integrateData.get(index)).set(3,Double.parseDouble(row.get(3).toString())+Double.parseDouble(((Vector)integrateData.get(index)).get(3).toString()));
                                    }

                                }
                            }catch (Exception err){
                                err.printStackTrace();
                            }

                        }else {
                            JOptionPane.showMessageDialog(mainPanel, "请选择Excel文件", "温馨提示",JOptionPane.WARNING_MESSAGE);
                        }
                    }
                    if(integrateData.size()!=0){
                        title=new Vector();
                        title.add("序号");
                        title.add("条目名称");
                        title.add("条目规格");
                        title.add("金额");
                        DefaultTableModel model=new DefaultTableModel(integrateData,title);//新建一个默认数据模型
                        excel.setModel(model);//显示汇总后的数据
                    }
                }
            }
        });

        fileChooser.addActionListener(new ActionListener() {//文件选择
            @Override
            public void actionPerformed(ActionEvent e) {
                JFileChooser jfc=new JFileChooser(new File("."));
                jfc.setFileSelectionMode(JFileChooser.FILES_ONLY  );
                jfc.showDialog(new JLabel(), "选择");
                File file=jfc.getSelectedFile();
                if(file!=null){
                    String filePath=file.getAbsolutePath();
                    String fileType = filePath.substring(filePath.lastIndexOf(".") + 1, filePath.length());
                    if(fileType.equals("xls")||fileType.equals("xlsx")){
                        initTable(file.getAbsolutePath());
                        fileName.setText("您打开了文件："+filePath);
                    }else {
                        JOptionPane.showMessageDialog(mainPanel, "请选择Excel文件", "温馨提示",JOptionPane.WARNING_MESSAGE);
                    }
                }
            }
        });
        arrange.addActionListener(new ActionListener() {//对数据进行整理
            @Override
            public void actionPerformed(ActionEvent e) {
                Vector arrangeData;//整理后的数据
                if(data==null){
                    JOptionPane.showMessageDialog(mainPanel, "请先选择文件", "温馨提示",JOptionPane.WARNING_MESSAGE);
                }else {
                    arrangeData=(Vector)data.clone();//克隆原始数据，在此基础上进行修改
                    for(int i=0;i<arrangeData.size();i++){
                        Vector row=(Vector)((Vector)arrangeData.get(i)).clone();
                        for(int j=4;j<7;j++){
                            row.set(j,"");
                        }
                        arrangeData.set(i,row);
                        if(row.get(1)==null || row.get(1).equals("")){
                            arrangeData.remove(i);
                        }
                    }
                    for(int i=0;i<data.size();i++){
                        String str;//结算表条目和规格的整合串
                        if(((Vector)data.get(i)).get(4)!=null && ((Vector)data.get(i)).get(5)!=null) {
                            str=((Vector)data.get(i)).get(4)+((Vector)data.get(i)).get(5).toString();
                        }
                        else {
                            str="";
                        }
                        if(map.get(str)!=null){
                            int index=Integer.parseInt(map.get(str).toString());
                            Vector row=(Vector)((Vector) data.get(index)).clone();
                            for(int j=4;j<7;j++){
                                row.set(j,((Vector)data.get(i)).get(j));
                            }
                            arrangeData.set(index,row);
                        }else {
                            Vector row=(Vector) ((Vector)data.get(i)).clone();
                            for(int j=0;j<4;j++){
                                row.set(j,"");
                            }
                            if(row.get(4)!=null && !row.get(4).equals("")){
                                row.set(0,"*");
                                arrangeData.add(row);
                            }
                        }
                    }
                    for(int i=0;i<arrangeData.size();i++){
                        Vector row=(Vector)((Vector)arrangeData.get(i)).clone();
                        if(!(row.get(3).equals("")&&row.get(6).equals(""))){
                            double leftValue,rightValue;//定义两次金额
                            if(row.get(3).equals("")){
                                leftValue=0;
                            }
                            else {
                                leftValue=Double.parseDouble(row.get(3).toString());
                            }
                            if(row.get(6).equals("")){
                                rightValue=0;
                            }else {
                                rightValue=Double.parseDouble(row.get(6).toString());
                            }
                            row.set(7,leftValue-rightValue);
                            arrangeData.set(i,row);
                        }else {
                            row.set(7,"");
                            arrangeData.set(i,row);
                        }
                    }
                    DefaultTableModel model=new DefaultTableModel(arrangeData,title);//新建一个默认数据模型
                    excel.setModel(model);//初始化数据
                }

            }
        });
        saveFile.addActionListener(new ActionListener() {//保存文件到指定路径
            @Override
            public void actionPerformed(ActionEvent e) {
                if(excel.getRowCount()==0){
                    JOptionPane.showMessageDialog(mainPanel, "没有可以保存的数据,请先选择文件", "温馨提示",JOptionPane.WARNING_MESSAGE);
                }else {
                    JFileChooser jfc=new JFileChooser(new File("."));
                    jfc.setFileSelectionMode(JFileChooser.FILES_ONLY );
                    jfc.showDialog(new JLabel(), "选择");
                    File file=jfc.getSelectedFile();
                    if (file!=null){
                        try {
                            WriteExcel writeExcel=new WriteExcel(file.getAbsolutePath());
                            Vector  row=new Vector();
                            for(int i=0;i<excel.getColumnModel().getColumnCount();i++){
                                row.add(excel.getColumnModel().getColumn(i).getHeaderValue());
                            }
                            writeExcel.write(row,0);
                            for(int i=0;i<excel.getRowCount();i++){
                                row=new Vector();
                                for (int j=0;j<excel.getColumnCount();j++){
                                    row.add(excel.getValueAt(i,j));
                                }
                                writeExcel.write(row,i+1);
                            }
                        }catch (Exception err){
                            err.printStackTrace();
                        }
                    }
                }

            }
        });
    }

    public void initTable(String filePath) {
        Vector row;//每行的元素
        int hashIndex=0;//hash索引的初始值
        data=new Vector();//所有行的元素
        map=new HashMap();//初始化hashmap
        title=new Vector();//表格头
        try {
            ReadExcel readExcel=new ReadExcel(filePath);
            if(!(readExcel.getAllRowNumber()<0)){
                title=readExcel.readLine(0);
            }
            for(int i=1;i<=readExcel.getAllRowNumber();i++){
                row=readExcel.readLine(i);
                if(row!=null){
                    if(!row.get(1).equals("")){
                        map.put(row.get(1)+row.get(2).toString(),hashIndex++);
                    }
                    data.add(row);
                }
            }
        }catch (Exception e){
            e.printStackTrace();
        }
        DefaultTableModel model=new DefaultTableModel(data,title);//新建一个默认数据模型
        excel.setModel(model);//初始化数据
    }

    public static void main(String[] args) {
        JFrame frame = new JFrame("ExcelTable");
        ExcelTable excelTable=new ExcelTable();
        frame.setContentPane(excelTable.mainPanel);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.pack();
        frame.setVisible(true);
    }

}
