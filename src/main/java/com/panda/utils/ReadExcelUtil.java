package com.panda.utils;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

public class ReadExcelUtil extends AbstractExcel{

    private final static Logger logger = LoggerFactory.getLogger(ReadExcelUtil.class);
    /**
     * 读取sheet表格
     * @param file 文件
     * @param startRowIndex 读取起始行 0是第一行
     * @return
     */
    public  List<List<String>> readSheetWithDOM(File file, int startRowIndex){
        Workbook workbook = loadExcel(file);
        Sheet sheet = workbook.getSheetAt(0);
        if(null == sheet){
            logger.warn("第一个sheet页不存在");
            return null;
        }
        int endRowIndex = sheet.getLastRowNum();
        if(logger.isDebugEnabled()){
            logger.debug("文件【{}】总计【{}】行",file.getName(),endRowIndex);
        }
        if(startRowIndex>endRowIndex){
            logger.error("读取起始行【{}】大于文件最大行【{}】",startRowIndex,endRowIndex);
            return null;
        }
        List<List<String>> data = new ArrayList<>(endRowIndex-startRowIndex);
        for(int rowIndex=startRowIndex; rowIndex<=endRowIndex; rowIndex++){
            List<String> rowData = new ArrayList<>();

        }

        return null;
    }

}
