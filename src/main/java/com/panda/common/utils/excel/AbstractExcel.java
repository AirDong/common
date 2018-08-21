package com.panda.common.utils.excel;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.*;

/**
 * Excel 抽象工具类
 */
public abstract class AbstractExcel {

    /**
     * 加载Excel
     * @param file
     * @return
     */
    public Workbook loadExcel(File file){
        if(file == null||!file.exists()){
            throw new RuntimeException("文件不存在");
        }
        InputStream ins=null;
        Workbook workbook=null;
        try {
            ins = new FileInputStream(file);
            workbook = WorkbookFactory.create(ins);
        }catch (InvalidFormatException e){
            throw new RuntimeException("文件格式不正确");
        }catch (Exception e){
            throw new RuntimeException(e);
        }
        return workbook;
    }

    /**
     * 加载Excel
     * @param filePath 文件路径
     * @return
     */
    public Workbook loadExcel(String filePath){
        if(filePath==null||filePath.length()==0){
            throw new RuntimeException("文件路径不正确");
        }
        File file = new File(filePath);
        return loadExcel(file);
    }

    /**
     * 关闭workBook对象
     * @param workbook
     */
    public void closeExcel(Workbook workbook){
        if(workbook!=null){
            try{
                workbook.close();
            }catch (Exception e){
                e.printStackTrace();
            }
        }
    }

    /**
     * 写入文件并关闭workbook对象
     * @param workbook
     * @param filePath
     * @param isDel 是否删除原文件
     */
    public void wirteToExcel(Workbook workbook,String filePath,boolean isDel)throws IOException {
        File file = new File(filePath);
        wirteToExcel(workbook,file,isDel);
    }
    /**
     * 写入文件并关闭workbook对象
     * @param workbook
     * @param file
     * @param isDel
     */
    public void wirteToExcel(Workbook workbook,File file,boolean isDel) throws IOException {
        OutputStream out = new FileOutputStream(file);
        workbook.write(out);
        closeExcel(workbook);
    }



}
