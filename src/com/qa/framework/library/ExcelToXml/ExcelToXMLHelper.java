package com.qa.framework.library.ExcelToXml;

import com.library.common.StringHelper;
import com.library.common.XmlHelper;
import com.qa.framework.exception.NoSuchNameInExcelException;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.dom4j.Element;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by Administrator on 2016/12/30.
 */
public class ExcelToXMLHelper {
    protected static Logger logger=Logger.getLogger(ExcelToXMLHelper.class);
    public static Map<String ,List<Map<String,String>>>  readExcel(File file) throws IOException,NoSuchNameInExcelException {
        if (file.getName().endsWith("xls")){
            return readExcelXLS(file);
        }else if (file.getName().endsWith("xlsx")){
            return readExcelXLSX(file);
        }else {
            return null;
        }
    }

    private static Map<String ,List<Map<String,String>>> readExcelXLSX(File file) throws IOException {
        XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(file));
       Map<String ,List<Map<String,String>>> excelValueMaps=new HashMap<String ,List<Map<String,String>>>();
        for (int numSheet=0;numSheet<wb.getNumberOfSheets();numSheet++){
            List<Map<String,String>> sheetValueList=new ArrayList<Map<String,String>>();
            String sheetName=wb.getSheetName(numSheet).trim().toLowerCase();
            sheetValueList=getSheetValues(wb.getSheetAt(numSheet));
            excelValueMaps.put(sheetName,sheetValueList);
        }
        return excelValueMaps;
    }

    private static List<Map<String, String>> getSheetValues(XSSFSheet sheet) {
        List<Map<String,String>> sheetValues=new ArrayList<Map<String, String>>();
        List<String> attributes=getCellValues(sheet.getRow(1));
        for (int rowNum=2;rowNum<=sheet.getLastRowNum();rowNum++){
            XSSFRow row=sheet.getRow(rowNum);
            if (row!=null) {
                Map<String, String> data = new HashMap<String, String>();
                int minCell = row.getFirstCellNum();
                int maxCell = row.getLastCellNum();
                for (int num = minCell; num < maxCell; num++) {
                    XSSFCell xssfCell = row.getCell(num);
                    if (xssfCell == null) {
                        data.put(attributes.get(num), "");
                    } else {
                        data.put(attributes.get(num), xssfCell.toString());
                    }

                }
                sheetValues.add(data);
            }
        }
        return sheetValues;
    }

    private static List<String> getCellValues(XSSFRow row) {
        List<String> cellValues=new ArrayList<String>();
        int minCell=row.getFirstCellNum();
        int maxCell=row.getLastCellNum();
        for (int num=minCell;num<maxCell;num++){
            XSSFCell xssfCell=row.getCell(num);
            cellValues.add(xssfCell.toString().toLowerCase());
        }
        return cellValues;
    }

    private static List<String> getSheetNames(XSSFWorkbook wb) {
        List<String> sheetNames=new ArrayList<String>();
        return sheetNames;
    }

    private static Map<String ,List<Map<String,String>>> readExcelXLS(File file) throws IOException,NoSuchNameInExcelException {
        InputStream is=new FileInputStream(file);
        HSSFWorkbook hssfWorkbook=new HSSFWorkbook(is);
        hssfWorkbook.getNumberOfSheets();
        hssfWorkbook.getSheet("");
        return null;
    }
    public static void processExcelValueToXml(String excelPath,String outputPath) throws IOException,NoSuchNameInExcelException {
        File file = new File(excelPath);
        Map<String ,List<Map<String,String>>> excelValueMaps=readExcel(file);
        List<Map<String,String>> dataConfigList=excelValueMaps.get("dataconfig");
        List<Map<String,String>> cookieList=excelValueMaps.get("cookie");
        List<Map<String,String>> cookieProcessList=excelValueMaps.get("cookieprocess");
        List<Map<String,String>> functionList=excelValueMaps.get("function");
        List<Map<String,String>> setupList=excelValueMaps.get("setup");
        List<Map<String,String>> paramList=excelValueMaps.get("param");
        List<Map<String,String>> testDataList=excelValueMaps.get("testdata");
        List<Map<String,String>> sqlList=excelValueMaps.get("sql");
        for (Map<String ,String> dataconfig :dataConfigList){
            String xmlName=dataconfig.get("name");
            List<Map<String,String>> testDataForSameUrl=new ArrayList<Map<String,String>>();
            for (Map<String,String> testdata:testDataList){
                if (testdata.get("dataconfig")!=""&&testdata.get("dataconfig").equals(xmlName)){
                    testDataForSameUrl.add(testdata);
                }
            }
            String url=dataconfig.get("url");
            String method=dataconfig.get("method");
            XmlHelper xml = new XmlHelper();
            xml.createDocument();
            Element root = xml.createDocumentRoot("DataConfig");
            xml.addAttribute(root,"url",url);
            xml.addAttribute(root,"httpMethod",method);
            for (Map<String,String> testDate:testDataForSameUrl){
                Element testdate=xml.addChildElement(root,"TestData");
                xml.addAttribute(testdate,"name",testDate.get("name"));
                xml.addAttribute(testdate,"desc",testDate.get("desc"));
                addCookieProcessAttr(xml,testdate,testDate,cookieProcessList);
                addBeforeChild(xml,testdate,testDate,functionList,sqlList);
                addHeaderChild(xml,testdate,testDate,cookieList);
                addSetupChild(xml,testdate,testDate,setupList,dataConfigList,paramList,cookieProcessList,sqlList,functionList);
                addParamChild(xml,testdate,testDate,paramList,sqlList,functionList);
                addExpectResult(xml,testdate);
                addAfterChild(xml,testdate,testDate,paramList,sqlList,functionList);
            }
            if (outputPath.endsWith("/")) {
                xml.saveTo(outputPath + xmlName + ".xml");
            } else {
                xml.saveTo(outputPath + File.separator + xmlName + ".xml");
            }
        }

    }

    private static void addExpectResult(XmlHelper xml, Element testdate) {
        Element expectResult=xml.addChildElement(testdate,"ExpectResult");
    }

    private  static void addAfterChild(XmlHelper xml, Element testdate, Map<String, String> testDate, List<Map<String, String>> paramList, List<Map<String, String>> sqlList, List<Map<String, String>> functionList) {
        if (testDate.get("before")!=null){
            Element after=xml.addChildElement(testdate,"After");
            String[] values=testDate.get("after").split(",");
            for (String value:values){
                Map<String,String> map= getMapByName(value,functionList,"Function");
                if(map.size()!=0){
                    Element function=xml.addChildElement(after,"Function");
                    xml.addAttribute(function,"ClassName",map.get("classname"));
                    xml.addAttribute(function,"MethodName",map.get("methodname"));
                }
            }
        }
    }

    private static void addCookieProcessAttr(XmlHelper xml, Element element, Map<String, String> map, List<Map<String, String>> cookieProcessList) {
        Map<String,String> cookieprocessMap=getMapByName(map.get("cookieprocess"),cookieProcessList,"CookieProcess");
        if (cookieprocessMap.size()!=0){
            xml.addAttribute(element,"storeCookie",cookieprocessMap.get("storecookie").toLowerCase());
            xml.addAttribute(element,"useCookie",cookieprocessMap.get("usecookie").toLowerCase());
        }
    }

    private static Map<String, String> getMapByName(String valueName, List<Map<String, String>> mapList, String mapName) {
        Map<String,String> setupMap= null;
        try {
            setupMap = isNameinMaps(valueName,mapList);
        } catch (NoSuchNameInExcelException e) {
            logger.info("----请检查"+valueName+"在"+mapName+"中是否存在----");
            e.printStackTrace();
        }
        return setupMap;
    }

    private static void addSetupChild(XmlHelper xml, Element testdate, Map<String, String> testDate, List<Map<String, String>> setupList, List<Map<String, String>> dataConfigList, List<Map<String, String>> paramList, List<Map<String, String>> cookieProcessList, List<Map<String, String>> sqlList, List<Map<String, String>> functionList) {
        if (testDate.get("setup")!=null){
            String[] values=testDate.get("setup").split(",");
            for (String value:values){
                Element setup=xml.addChildElement(testdate,"Setup");
                Map<String,String> setupMap=getMapByName(value,setupList,"Setup");
                if(setupMap.size()!=0){
                    xml.addAttribute(setup,"name",setupMap.get("name"));
                }
                Map<String,String> map2=getMapByName(setupMap.get("dataconfig"),dataConfigList,"DataConfig");
                if (map2!=null){
                    xml.addAttribute(setup,"url",map2.get("url"));
                    xml.addAttribute(setup,"httpMothd",map2.get("httpmethod"));
                }
                addCookieProcessAttr(xml,setup,setupMap,cookieProcessList);
                addParamChild(xml,setup,setupMap,paramList,sqlList,functionList);
            }

        }
    }

    private static void addParamChild(XmlHelper xml, Element element, Map<String, String> map, List<Map<String, String>> paramList, List<Map<String, String>> sqlList, List<Map<String, String>> functionList) {
        if (map.get("param")!=null){
            String[] values=map.get("param").split(",");
            for (String value:values){
                Element param=xml.addChildElement(element,"Param");
                Map<String,String> paramMap=getMapByName(value,paramList,"Param");
                if (paramMap.size()!=0) {
                    xml.addAttribute(param,"name",paramMap.get("key"));
                    xml.addAttribute(param,"value",paramMap.get("value"));
                    if (paramMap.get("value").contains("#{")){
                        List<String> lists = StringHelper.find(paramMap.get("value"), "#\\{[a-zA-Z0-9._]*\\}");
                        for (String list : lists) {
                            String sqlName = list.substring(2, list.length() - 1).split("\\.")[0];
                            Map<String,String> sqlMap=getMapByName(sqlName,sqlList,"Sql");
                            Element sql=xml.addChildElement(param,"Sql");
                            xml.addAttribute(sql,"name",sqlName);
                            xml.setText(sql,sqlMap.get("value"));
                        }
                    }
                }
            }
        }
    }


    private static void addHeaderChild(XmlHelper xml, Element testdate, Map<String, String> testDate, List<Map<String, String>> cookieList) {
        if (testDate.get("before")!=null){
            Element header=xml.addChildElement(testdate,"Header");
            String[] values=testDate.get("header").split(",");
            for (String value:values){
                Map<String,String> map=getMapByName(value,cookieList,"Cookie");
                if(map!=null){
                    Element cookie=xml.addChildElement(header,"Cookie");
                    xml.addAttribute(cookie,"name",map.get("key"));
                    xml.addAttribute(cookie,"value",map.get("value"));
                }
            }
        }
    }

    private static void addBeforeChild(XmlHelper xml, Element testdate, Map<String, String> testDate, List<Map<String, String>> functionList, List<Map<String, String>> sqlList) {
        if (testDate.get("before")!=null){
            Element before=xml.addChildElement(testdate,"Before");
            String[] values=testDate.get("before").split(",");
            for (String value:values){
                Map<String,String> map= null;
                try {
                    map = isNameinMaps(value,functionList);
                } catch (NoSuchNameInExcelException e) {

                    e.printStackTrace();
                }
                if(map.size()!=0){
                     Element function=xml.addChildElement(before,"Function");
                     xml.addAttribute(function,"ClassName",map.get("classname"));
                     xml.addAttribute(function,"MethodName",map.get("methodname"));
                 }
            }
        }
    }

    private static  Map<String,String> isNameinMaps(String value, List<Map<String, String>> mapList) throws NoSuchNameInExcelException {
        Map<String,String> inMap=new HashMap<String,String>();
        for (Map<String,String> map:mapList){
            String mapName=map.get("name").trim();
            if (mapName.equalsIgnoreCase(value)){
                inMap=map;
                break;
            }
        }
        if (inMap.size()==0){
            throw new NoSuchNameInExcelException(value);
        }
        return inMap;
    }

}
