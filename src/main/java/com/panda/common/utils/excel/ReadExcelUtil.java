package com.panda.common.utils.excel;

import com.monitorjbl.xlsx.StreamingReader;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.List;

/**
 * 读取Excel 内容工具类
 */
public class ReadExcelUtil extends AbstractExcel{

    private final static Logger logger = LoggerFactory.getLogger(ReadExcelUtil.class);
    /**
     * 读取sheet表格
     * @param file 文件
     * @param startRowIndex 读取起始行 0是第一行
     * @return  List<List<String>>  List<String> 第一列为行号
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
            //记录所在行数
            rowData.add(String.valueOf(rowIndex));
            Row row = sheet.getRow(rowIndex);
            for(int cellIndex=0;cellIndex<row.getLastCellNum();cellIndex++){
                Cell cell = row.getCell(cellIndex);
                if(null == cell){
                    rowData.add(null);
                    continue;
                }
                CellType cellType = cell.getCellTypeEnum();
                if(cellType == CellType.ERROR){
                    rowData.add(null);
                }else if(cellType == CellType.BLANK){
                    rowData.add(null);
                }else if(cellType == CellType.NUMERIC){
                    DecimalFormat df = new DecimalFormat("0.####################");
                    rowData.add(df.format(cell.getNumericCellValue()));
                }else {
                    String val = cell.getStringCellValue();
                    rowData.add(val==null?null:val.trim());
                }
            }
            data.add(rowData);
        }
        return data;
    }

    /**
     * 读取07版本上的大数据文件
     * @param file
     * @return
     * @throws IOException
     */
    public static List<List<String>> readXlsxSheetWithSax(File file)throws IOException{
        List<List<String>> data = new ArrayList<>(512);
        FileInputStream in = new FileInputStream(file);
        Workbook wk = StreamingReader.builder()
                //缓存到内存中的行数，默认是10
                .rowCacheSize(100)
                ////读取资源时，缓存到内存的字节大小，默认是1024
                .bufferSize(4096)
                // //打开资源，必须，可以是InputStream或者是File，注意：只能打开XLSX格式的文件
                .open(in);
        Sheet sheet = wk.getSheetAt(0);
        for (Row row : sheet) {
            List<String> rowData = new ArrayList<>();
            rowData.add(row.getRowNum()+"");
            //遍历所有的列
            for (Cell cell : row) {
                rowData.add(cell.getStringCellValue());
            }
            data.add(rowData);
        }
        return data;
    }


}
