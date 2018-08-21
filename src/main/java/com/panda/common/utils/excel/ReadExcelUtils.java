package com.panda.common.utils.excel;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.*;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.*;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * 基于XSSF and SAX (Event API)
 * 读取excel的第一个Sheet的内容
 *
 */
public class ReadExcelUtils {
    private int headCount = 0;
    private List<List<String>> list = new ArrayList<List<String>>();

    /**
     * 采用DOM的形式进行解析
     * @param file
     * @param headRowCount   跳过读取的表头的行数
     * @return
     * @throws IOException
     * @throws InvalidFormatException
     * @throws Exception
     */
    public  List<List<String>> processDOMReadSheet(File file,int headRowCount) throws InvalidFormatException, IOException {
        InputStream ins=null;
        Workbook workbook=null;
        try {
            ins=new FileInputStream(file);
            workbook = WorkbookFactory.create(ins);
            return this.processDOMRead(workbook, headRowCount);
        }finally {
            if(workbook!=null){
                try {
                    workbook.close();
                }catch (Exception e){
                    e.printStackTrace();
                }
            }
            if(ins!=null){
                try {
                    ins.close();
                }catch (Exception e){
                    e.printStackTrace();
                }
            }
        }
    }

    public  void write(List<List<String>> list,File file) throws InvalidFormatException, IOException {
        InputStream ins=null;
        Workbook workbook=null;
        ins=new FileInputStream(file);
        workbook = WorkbookFactory.create(ins);
        Sheet sheet = workbook.getSheetAt(0);
        for(List<String> rowData:list){
            int rowIndex = Integer.parseInt(rowData.get(0));
            String webgateStatus = rowData.get(1);
            String webgatePath = rowData.get(2);
            String mapiStatus = rowData.get(3);
            String mapiPath = rowData.get(4);
            String orderId = rowData.get(5);

            Row row = sheet.getRow(rowIndex);
           String str = row.getCell(0).getStringCellValue();
           if(!orderId.equals(str)){
               throw new RuntimeException(orderId+"不匹配");
           }
            Cell cell9 = row.createCell(9);
            cell9.setCellValue(webgateStatus);
            Cell cell10 = row.createCell(10);
            cell10.setCellValue(mapiStatus);
            Cell cell11 = row.createCell(11);
            cell11.setCellValue(webgatePath);
            Cell cell12 = row.createCell(12);
            cell12.setCellValue(mapiPath);
        }
        OutputStream out = new FileOutputStream(new File("D://12.xlsx"));
        workbook.write(out);
        workbook.close();
    }

    /**
     * 采用SAX进行解析xlxs
     * @param filename
     * @param headRowCount
     * @return
     * @throws OpenXML4JException
     * @throws IOException
     * @throws SAXException
     * @throws Exception
     */
    public List<List<String>> processSAXReadSheet(String filename,int headRowCount) throws IOException, OpenXML4JException, SAXException   {
        headCount = headRowCount;
        OPCPackage pkg = OPCPackage.open(filename);
        XSSFReader r = new XSSFReader( pkg );
        SharedStringsTable sst = r.getSharedStringsTable();
        XMLReader parser = fetchSheetParser(sst);

        Iterator<InputStream> sheets = r.getSheetsData();
        InputStream sheet = sheets.next();
        InputSource sheetSource = new InputSource(sheet);
        parser.parse(sheetSource);
        sheet.close();

        return list;
    }

    private XMLReader fetchSheetParser(SharedStringsTable sst) throws SAXException {
        XMLReader parser =
                XMLReaderFactory.createXMLReader(
                        "org.apache.xerces.parsers.SAXParser"
                );
        ContentHandler handler = new SheetHandler(sst);
        parser.setContentHandler(handler);
        return parser;
    }

    /**
     * SAX 解析excel
     */
    private class SheetHandler extends DefaultHandler {
        private SharedStringsTable sst;
        private String lastContents;
        private boolean nextIsString;
        private boolean isNullCell;
        //读取行的索引
        private int rowIndex = 0;
        //是否重新开始了一行
        private boolean curRow = false;
        private List<String> rowContent;

        private SheetHandler(SharedStringsTable sst) {
            this.sst = sst;
        }

        public void startElement(String uri, String localName, String name,
                                 Attributes attributes) throws SAXException {
            //节点的类型
            //System.out.println("---------begin:" + name);
            if(name.equals("row")){
                rowIndex++;
            }
            //表头的行直接跳过
            if(rowIndex > headCount){
                curRow = true;
                // c => cell
                if(name.equals("c")) {
                    String cellType = attributes.getValue("t");
                    if(null == cellType){
                        isNullCell = true;
                    }else{
                        if(cellType.equals("s")) {
                            nextIsString = true;
                        } else {
                            nextIsString = false;
                        }
                        isNullCell = false;
                    }
                }
                // Clear contents cache
                lastContents = "";
            }
        }

        public void endElement(String uri, String localName, String name)
                throws SAXException {
            //System.out.println("-------end："+name);
            if(rowIndex > headCount){
                if(nextIsString) {
                    int idx = Integer.parseInt(lastContents);
                    lastContents = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
                    nextIsString = false;
                }
                if(name.equals("v")) {
                    //System.out.println(lastContents);
                    if(curRow){
                        //是新行则new一行的对象来保存一行的值
                        if(null==rowContent){
                            rowContent = new ArrayList<String>();
                        }
                        rowContent.add(lastContents);
                    }
                }else if(name.equals("c") && isNullCell){
                    if(curRow){
                        //是新行则new一行的对象来保存一行的值
                        if(null==rowContent){
                            rowContent = new ArrayList<String>();
                        }
                        rowContent.add(null);
                    }
                }

                isNullCell = false;

                if("row".equals(name)){
                    list.add(rowContent);
                    curRow = false;
                    rowContent = null;
                }
            }

        }

        public void characters(char[] ch, int start, int length)
                throws SAXException {
            lastContents += new String(ch, start, length);
        }
    }

    /**
     * DOM的形式解析execl
     * @param workbook
     * @param headRowCount
     * @return
     * @throws InvalidFormatException
     * @throws IOException
     */
    private List<List<String>> processDOMRead(Workbook workbook, int headRowCount) throws InvalidFormatException, IOException {
        headCount = headRowCount;
        Sheet sheet = workbook.getSheetAt(0);
        //行数
        int endRowIndex = sheet.getLastRowNum();
        Row row = null;
        List<String> rowList = null;
        for(int i=headCount; i<=endRowIndex; i++){
            rowList = new ArrayList<String>();
            row = sheet.getRow(i);
            rowList.add(i+"");
            for(int j=0; j<row.getLastCellNum();j++){
                if(null==row.getCell(j)){
                    rowList.add(null);
                    continue;
                }
                int dataType = row.getCell(j).getCellType();
                if(dataType == Cell.CELL_TYPE_NUMERIC){
                    DecimalFormat df = new DecimalFormat("0.####################");
                    rowList.add(df.format(row.getCell(j).getNumericCellValue()));
                }else if(dataType == Cell.CELL_TYPE_BLANK){
                    rowList.add(null);
                }else if(dataType == Cell.CELL_TYPE_ERROR){
                    rowList.add(null);
                }else{
                    //这里的去空格根据自己的情况判断
                    String valString = row.getCell(j).getStringCellValue();
                    if(null!=valString){
                        valString=valString.trim();
                    }
                    rowList.add(valString);
                }
            }
            list.add(rowList);
        }
        return list;
    }


    public static void main(String[] args)throws Exception{
        ReadExcelUtils utils = new ReadExcelUtils();
        List<List<String>> list = utils.processDOMReadSheet(new File("D://11.xlsx"),1);

        List<List<String>> result = new ArrayList<>();

       /* String webgateStatus = rowData.get(1);
        String webgatePath = rowData.get(2);
        String mapiStatus = rowData.get(3);
        String mapiPath = rowData.get(4);*/
        for(List<String> orderlist:list){
            String orderId = orderlist.get(1);
            String indexRow = orderlist.get(0);
            List<String> rowData = new ArrayList<>();
            rowData.add(indexRow);
            if(searchFile(orderId+".txt","D:\\ipay\\webgate11")){
                rowData.add("是");
                rowData.add("生产一套");
            }else if(searchFile(orderId+".txt","D:\\ipay\\webgate12")){
                rowData.add("是");
                rowData.add("生产二套");
            }else{
                if(isExist("D:\\ipay\\webgate11\\"+orderId+".txt")||isExist("D:\\ipay\\webgate12\\"+orderId+".txt")){
                    rowData.add("否");
                    rowData.add("");
                }else {
                    rowData.add("无日志");
                    rowData.add("");
                }
            }
            if(searchFile(orderId+".txt","D:\\ipay\\mapi11")){
                rowData.add("是");
                rowData.add("生产一套");
            }else if(searchFile(orderId+".txt","D:\\ipay\\mapi12")){
                rowData.add("是");
                rowData.add("生产二套");
            }else{
                if(isExist("D:\\ipay\\mapi11\\"+orderId+".txt")||isExist("D:\\ipay\\mapi12\\"+orderId+".txt")){
                    rowData.add("否");
                    rowData.add("");
                }else {
                    rowData.add("无日志");
                    rowData.add("");
                }

            }
            rowData.add(orderId);
            result.add(rowData);
        }
        utils.write(result,new File("D://11.xlsx"));

       /* String[] webgate11s = new String[]{"2018-08-02",
                "2018-08-03",
                "2018-08-06",
                "2018-08-07",
                "2018-08-08",
                "2018-08-09",
                "2018-08-10",
                "2018-08-11",
                "2018-08-13",
                "2018-08-14",
                "2018-08-15",
                "2018-08-16",
                "2018-08-17"};
        String[] webgate12s = new String[]{
                "2018-08-02",
                "2018-08-03",
                "2018-08-04",
                "2018-08-05",
                "2018-08-06",
                "2018-08-07",
                "2018-08-08",
                "2018-08-09",
                "2018-08-11",
                "2018-08-12",
                "2018-08-13",
                "2018-08-14",
                "2018-08-15",
                "2018-08-16",
                "2018-08-17"
        };

        String[] mapi12 = new String[]{
                "2018-08-01",
                "2018-08-01",
                "2018-08-01",
                "2018-08-01",
                "2018-08-01",
                "2018-08-01",
                "2018-08-02",
                "2018-08-02",
                "2018-08-02",
                "2018-08-02",
                "2018-08-02",
                "2018-08-02",
                "2018-08-01",
                "2018-08-02",
                "2018-08-04",
                "2018-08-04",
                "2018-08-04",
                "2018-08-04",
                "2018-08-04",
                "2018-08-04",
                "2018-08-04",
                "2018-08-05",
                "2018-08-05",
                "2018-08-05",
                "2018-08-05",
                "2018-08-05",
                "2018-08-05",
                "2018-08-05",
                "2018-08-06",
                "2018-08-06",
                "2018-08-06",
                "2018-08-06",
                "2018-08-06",
                "2018-08-06",
                "2018-08-06",
                "2018-08-07",
                "2018-08-07",
                "2018-08-07",
                "2018-08-03",
                "2018-08-07",
                "2018-08-07",
                "2018-08-07",
                "2018-08-11",
                "2018-08-12",
                "2018-08-12",
                "2018-08-12",
                "2018-08-12",
                "2018-08-13",
                "2018-08-13",
                "2018-08-13",
                "2018-08-13",
                "2018-08-14",
                "2018-08-14",
                "2018-08-14",
                "2018-08-15",
                "2018-08-11",
                "2018-08-15",
                "2018-08-15",
                "2018-08-16",
                "2018-08-16",
                "2018-08-16",
                "2018-08-16",
                "2018-08-17",
                "2018-08-15",
                "2018-08-17"
        };

        String[] mapi11 = new String[]{
                "2018-08-01",
                "2018-08-02",
                "2018-08-03",
                "2018-08-04",
                "2018-08-05",
                "2018-08-06",
                "2018-08-07",
                "2018-08-08",
                "2018-08-09",
                "2018-08-10",
                "2018-08-11",
                "2018-08-12",
                "2018-08-13",
                "2018-08-16"
        };

        ReadExcelUtils utils = new ReadExcelUtils();
        List<List<String>> list = utils.processDOMReadSheet(new File("D://11.xlsx"),1);
        //
        File webgate11 =new File("D://webgate11.txt");
        webgate11.createNewFile();
        FileWriter  out=new FileWriter (webgate11);
        BufferedWriter webgate11bw= new BufferedWriter(out);
        //
        File webgate12 =new File("D://webgate12.txt");
        webgate12.createNewFile();
        FileWriter  out12=new FileWriter (webgate12);
        BufferedWriter webgate12bw= new BufferedWriter(out12);

        File mapi11File =new File("D://mapi11.txt");
        mapi11File.createNewFile();
        FileWriter  mapi11out=new FileWriter (mapi11File);
        BufferedWriter mapi11bw= new BufferedWriter(mapi11out);

        File mapi112File =new File("D://mapi12.txt");
        mapi112File.createNewFile();
        FileWriter  mapi12out=new FileWriter (mapi112File);
        BufferedWriter mapi12bw= new BufferedWriter(mapi12out);
        int i=0;
        for(List<String> arrStrs:list){
            // 订单号
            String orderId = arrStrs.get(0);
            //所在日志的日期
            String recodeDate = getDate(arrStrs.get(8));
            //
           *//*if(contian(webgate11s,recodeDate)){
               webgate11bw.write("grep '"+orderId+"' `find ./ -name 'catalina."+recodeDate+".out'` > /tmp/webgate11/"+orderId+".txt");
               webgate11bw.newLine();
           }
            if(contian(webgate12s,recodeDate)){
                webgate12bw.write("grep '"+orderId+"' `find ./ -name 'catalina."+recodeDate+".out'` > /tmp/webgate12/"+orderId+".txt");
                webgate12bw.newLine();
            }
            if(contian(mapi11,recodeDate)){
                mapi11bw.write("grep '"+orderId+"' `find ./ -name 'log-"+recodeDate+"-*.log'` > /tmp/mapi11/"+orderId+".txt");
                mapi11bw.newLine();
            }*//*
            if(contian(mapi12,recodeDate)){

                mapi12bw.write("grep '"+orderId+"' `find /ipaylogs/mop-mapi12/logs/mop-mapi/mop-mapi-2018-08 -name 'log-"+recodeDate+"-*.log'` > /tmp/mapi12/"+orderId+".txt");
                mapi12bw.newLine();
                mapi12bw.write("echo "+(++i)+"");
                mapi12bw.newLine();
            }
        }
        webgate11bw.close();webgate12bw.close(); mapi11bw.close();mapi12bw.close();

*/



    }


    private static String getDate(String dateStr){
        try {
            SimpleDateFormat sDateFormat=new SimpleDateFormat("yyyy-MM-dd HH:mm:ss"); //加上时间
            Date date = sDateFormat.parse(dateStr);
            Calendar calendar = Calendar.getInstance();
            calendar.setTime(date);
            calendar.add(Calendar.HOUR_OF_DAY,8);
            SimpleDateFormat sbf = new SimpleDateFormat("yyyy-MM-dd");
            return sbf.format(calendar.getTime());
        }catch (Exception e){
            e.printStackTrace();
        }
        return null;
    }

    public static boolean contian(String[] arrs,String str){
        for(String arr:arrs){
            if(str.equals(arr)){
                return true;
            }
        }
        return false;
    }


    public static boolean searchFile(String fileName,String filePath){
        File file = new File(filePath+"/"+fileName);
        if(file.exists()){//存在
            if(file.length()>0){
                return true;
            }
        }
        return false;
    }

    public static boolean isExist(String path){
        File file = new File(path);
        if(file.exists()){
            return true;
        }
        return false;
    }
}