package com.moan.hoe.base.util;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFFont;

import javax.servlet.http.HttpServletResponse;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelUtil {

    /**
    * @Description: 获取导入excel数据
    * @Param: [workbook, columnNum]
    * @return: java.util.List<java.util.List<java.lang.String>>
    * @Author: yym
    * @Date: 2019-12-15
    */
    public static List<List<String>> getImportExcelDatas(Workbook workbook,int columnNum){
        return ExcelUtil.getImportExcelDatas(workbook,0,columnNum);
    }

    /**
     * 获取导入excel数据
     * @param workbook
     * @param columnNum
     * @return
     */
    public static List<List<String>> getImportExcelDatas(Workbook workbook,int sheetIndex,int columnNum){
        List<List<String>> results = new ArrayList<List<String>>();
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        for(int i=0;sheet.getRow(i) != null;i++){
            Row row = sheet.getRow(i);
            List<String> result = new ArrayList<String>();
            boolean isEnd = true;
            for(int j=0;j<columnNum;j++){
                Cell cell = row.getCell(j);
                if(null != cell){
                    switch (cell.getCellType()){
                        case STRING:
                            result.add(cell.getStringCellValue());
                            break;
                        case NUMERIC:
                            if(DateUtil.isCellDateFormatted(cell)){
                                result.add(cell.getDateCellValue()+"");
                            }else {
                                result.add(cell.getNumericCellValue()+"");
                            }
                            break;
                        case BOOLEAN:
                            result.add(cell.getBooleanCellValue()+"");
                            break;
                        case FORMULA:
                            result.add(cell.getCellFormula());
                            break;
                        default:
                            result.add(null);
                            break;
                    }
                    if(UtilTools.isNotEmpty(result.get(j))){
                        isEnd = false;
                    }
                }else{
                    result.add(null);
                }
            }
            if(isEnd){
                break;
            }
            results.add(result);
        }
        return results;
    }

    /**
    * @Description: 获取excel指定sheet、row、cell 的值
    * @Param: [workbook, sheetIndex, rowIndex, colIndex]
    * @return: java.lang.String
    * @Author: yym
    * @Date: 2019-8-27
    */
    public static String getCellValue(Workbook workbook,int sheetIndex, int rowIndex, int colIndex){
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        if(sheet == null){
            return null;
        }
        Row row = sheet.getRow(rowIndex);
        if(row == null){
            return null;
        }
        Cell cell = row.getCell(colIndex);
        if(cell == null){
            return null;
        }
        String value = null;
        switch (cell.getCellType()){
            case STRING:
                value = cell.getStringCellValue();
                break;
            case NUMERIC:
                if(DateUtil.isCellDateFormatted(cell)){
                    value = cell.getDateCellValue()+"";
                }else {
                    value = cell.getNumericCellValue()+"";
                }
                break;
            case BOOLEAN:
                value = cell.getBooleanCellValue()+"";
                break;
            case FORMULA:
                value = cell.getCellFormula()+"";
                break;
        }
        return value;
    }

    /**
     * 导出excel
     * @param columnTitles
     * @param datas
     * @param fileName
     * @param columnWidth
     * @param columnWidthList
     * @param response
     */
    public static void createExportExcel(List<String> columnTitles, List<List<String>> datas,String fileName,
                                         String columnWidth,List<Integer> columnWidthList,
                                         HttpServletResponse response){
        Map<String,Object> info = new HashMap<>();
        info.put("columnTitles",columnTitles);
        info.put("datas",datas);
        info.put("fileName",fileName);
        if(null != columnWidth && !"".equals(columnWidth)){
            info.put("columnWidth",columnWidth);
        }
        if(null != columnWidthList && columnWidthList.size() != 0){
            info.put("columnWidthList",columnWidthList);
        }
        info.put("response",response);
        createExportExcel(info);
    }
    /**
     * 创建导出excel
     * @param info
     * @return
     */
    public static void createExportExcel(Map<String,Object> info){
        try {
            HttpServletResponse response = (HttpServletResponse) info.get("response");
            String fileName = (String)info.get("fileName");
            List<String> columnTitles = (List<String>)info.get("columnTitles");
            List<List<String>> datas = (List<List<String>>) info.get("datas");
            int columnWidth = 16;
            if(info.containsKey("columnWidth")){
                columnWidth = (Integer)info.get("columnWidth");
            }
            List<Integer> columnWidthList = (List<Integer>)info.get("columnWidthList");//各列列宽
            short rowHeight = 400;
            if(info.containsKey("rowHeight")){//行高
                rowHeight = Short.parseShort(info.get("rowHeight").toString());
            }
            OutputStream os = null;
            FileOutputStream fos = null;
            if(response != null){
                os = response.getOutputStream();// 取得输出流
                response.reset();// 清空输出流
                response.setHeader("Content-disposition", "attachment; filename="
                        + new String((fileName + CommonUtil.generateTimestampNo()).getBytes("GB2312"), "8859_1")
                        + ".xls");// 设定输出文件头
                response.setContentType("application/msexcel");// 定义输出类型
            }else {
                fos = (FileOutputStream)info.get("outputStream");
            }
            Workbook workbook = new HSSFWorkbook();
            CellStyle cellStyle = workbook.createCellStyle();
            cellStyle.setWrapText(true);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            Sheet sheet = workbook.createSheet();
            sheet.setDefaultColumnWidth(columnWidth);
            if(columnWidthList != null && columnWidthList.size() == columnTitles.size()){
                for(int i=0;i<columnWidthList.size();i++){
                    sheet.setColumnWidth(i,columnWidthList.get(i)*256);
                }
            }
            int rowIndex = 0;
            //导出模板 抬头
            if (info.containsKey("titleName")) {
                Row row = sheet.createRow(rowIndex);
                row.setHeight(rowHeight);
                row.createCell(0).setCellValue(info.get("titleName").toString());
                CellRangeAddress cellTitle = new CellRangeAddress(rowIndex, rowIndex, 0, columnTitles.size() -1);
                sheet.addMergedRegion(cellTitle);
                rowIndex++;
            }
            //导出说明
            if (info.containsKey("importDescription")) {
                Row row = sheet.createRow(rowIndex);
                row.setHeight(rowHeight);
                row.createCell(0).setCellValue(info.get("importDescription").toString());
                CellRangeAddress cellTitle = new CellRangeAddress(rowIndex, rowIndex, 0, columnTitles.size() -1);
                sheet.addMergedRegion(cellTitle);
                rowIndex++;
            }
            Row row = sheet.createRow(rowIndex);
            row.setHeight(rowHeight);
            for(int i=0;i<columnTitles.size();i++){
                row.createCell(i).setCellValue(columnTitles.get(i));
            }
            rowIndex++;
            for(int i=0;i<datas.size();i++){
                List<String> data = datas.get(i);
                row = sheet.createRow(rowIndex);
                row.setHeight(rowHeight);
                for(int j=0;j<data.size();j++){
                    Cell cell = row.createCell(j);
                    cell.setCellStyle(cellStyle);
                    cell.setCellValue(data.get(j));
                }
                rowIndex++;
            }
            if(os != null){
                workbook.write(os);
                os.flush();
                os.close();
            }else {
                workbook.write(fos);
                fos.flush();
                fos.close();
            }
        }catch (Exception e){
            e.printStackTrace();
        }
    }

    public static void createMultipleExportExcel(Map<String,Object> info){
        try {
            HttpServletResponse response = (HttpServletResponse) info.get("response");
            String fileName = (String)info.get("fileName");
            List<Map<String,Object>> dataList = (List<Map<String,Object>>) info.get("dataList");
            int columnWidth = 16;
            if(info.containsKey("columnWidth")){
                columnWidth = (Integer)info.get("columnWidth");
            }
            Boolean split = true;
            if(info.get("noSplit") != null && info.get("noSplit").toString().equals("1")){
                split = false;
            }
            OutputStream os = null;
            FileOutputStream fos = null;
            if(response != null){
                os = response.getOutputStream();// 取得输出流
                response.reset();// 清空输出流
                response.setHeader("Content-disposition", "attachment; filename="
                        + new String((fileName + CommonUtil.generateTimestampNo()).getBytes("GB2312"), "8859_1")
                        + ".xls");// 设定输出文件头
                response.setContentType("application/msexcel");// 定义输出类型
            }else {
                fos = (FileOutputStream)info.get("outputStream");
            }
            Workbook workbook = new HSSFWorkbook();
            CellStyle normalCellStyle = workbook.createCellStyle();
            normalCellStyle.setAlignment(HorizontalAlignment.CENTER);
            normalCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

            Map<String,Object> dataMap = null;
            for (int j=0;j<dataList.size();j++){
                dataMap = dataList.get(j);
                List<String> columnTitles = (List<String>)dataMap.get("columnTitles");
                List<List<String>> datas = (List<List<String>>) dataMap.get("datas");
                String sheetName = (String)dataMap.get("sheetName");
                Sheet sheet = workbook.createSheet();
                sheet.setDefaultColumnWidth(columnWidth);
                if(!UtilTools.isEmpty(sheetName)){
                    sheetName = sheetName.replaceAll("/","");
                    sheetName = sheetName.replaceAll("\\\\","");
                    workbook.setSheetName(j,sheetName);
                }
                int rowIndex = 0;
                Row row = sheet.createRow(rowIndex);
                boolean spanRow = false;
                List<Integer> indexList = new ArrayList<>();
                int colIndex = 0;
                Row secondRow = null;
                for(int i=0;i<columnTitles.size();i++){
                    if(columnTitles.get(i).indexOf(",") > 0){
                        spanRow = true;
                        String[] titles = columnTitles.get(i).split(",");
                        for(int t = 0;t < titles.length;t++){
                            if(t == 0){
                                Cell cell = row.createCell(colIndex);
                                cell.setCellValue(titles[t]);
                                cell.setCellStyle(normalCellStyle);
                                sheet.addMergedRegion(new CellRangeAddress(rowIndex,rowIndex,colIndex,colIndex+titles.length-1));
                                if(secondRow == null){
                                    secondRow = sheet.createRow(rowIndex+1);
                                }
                                cell = secondRow.createCell(colIndex);
                                cell.setCellValue("小项总分");
                                cell.setCellStyle(normalCellStyle);
                                colIndex++;
                            }else{
                                Cell cell = secondRow.createCell(colIndex++);
                                cell.setCellValue(titles[t]);
                                cell.setCellStyle(normalCellStyle);
                            }
                        }
                    }else{
                        indexList.add(colIndex);
                        Cell cell = row.createCell(colIndex++);
                        cell.setCellValue(columnTitles.get(i));
                        cell.setCellStyle(normalCellStyle);
                    }
                }
                if(spanRow){
                    for(Integer index : indexList){
                        sheet.addMergedRegion(new CellRangeAddress(rowIndex,rowIndex+1,index,index));
                    }
                    rowIndex++;
                }
                rowIndex++;
                for(int i=0;i<datas.size();i++){
                    List<String> data = datas.get(i);
                    row = sheet.createRow(rowIndex);
                    colIndex = 0;
                    for(int k=0;k<data.size();k++){
                        if(split && data.get(k).indexOf(",") > 0){
                            String[] scores = data.get(k).split(",");
                            for(int t = 0;t < scores.length;t++){
                                row.createCell(colIndex++).setCellValue(scores[t]);
                            }
                        }else{
                            row.createCell(colIndex++).setCellValue(data.get(k));
                        }
                    }
                    rowIndex++;
                }
            }
            if(os != null){
                workbook.write(os);
                os.flush();
                os.close();
            }else {
                workbook.write(fos);
                fos.flush();
                fos.close();
            }
        }catch (Exception e){
            e.printStackTrace();
        }
    }

    public static void createDetailsExportExcel(Map<String,Object> info){
        try {
            HttpServletResponse response = (HttpServletResponse) info.get("response");
            String fileName = (String)info.get("fileName");
            List<Map<String,Object>> dataList = (List<Map<String,Object>>) info.get("dataList");
            int columnWidth = 16;
            if(info.containsKey("columnWidth")){
                columnWidth = (Integer)info.get("columnWidth");
            }
            Boolean split = true;
            if(info.get("noSplit") != null && info.get("noSplit").toString().equals("1")){
                split = false;
            }
            OutputStream os = null;
            FileOutputStream fos = null;
            if(response != null){
                os = response.getOutputStream();// 取得输出流
                response.reset();// 清空输出流
                response.setHeader("Content-disposition", "attachment; filename="
                        + new String((fileName + CommonUtil.generateTimestampNo()).getBytes("GB2312"), "8859_1")
                        + ".xlsx");// 设定输出文件头
                response.setContentType("application/msexcel");// 定义输出类型
            }else {
                fos = (FileOutputStream)info.get("outputStream");
            }
            Workbook workbook = new HSSFWorkbook();
            CellStyle normalCellStyle = workbook.createCellStyle();
            normalCellStyle.setAlignment(HorizontalAlignment.CENTER);
            normalCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

            Map<String,Object> dataMap = null;
            for (int j=0;j<dataList.size();j++){
                dataMap = dataList.get(j);
                List<String> columnTitles = (List<String>)dataMap.get("columnTitles");
                List<List<String>> datas = (List<List<String>>) dataMap.get("datas");
                String sheetName = (String)dataMap.get("sheetName");
                Sheet sheet = workbook.createSheet();
                Row rowOne = sheet.createRow(0);
                Cell cellOne = rowOne.createCell(0);
                sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, datas.get(0).size()));
                cellOne.setCellValue(sheetName);
                cellOne.setCellStyle(normalCellStyle);
                sheet.setDefaultColumnWidth(columnWidth);
//                if(!UtilTools.isEmpty(sheetName)){
//                    sheetName = sheetName.replaceAll("/","");
//                    sheetName = sheetName.replaceAll("\\\\","");
//                    workbook.setSheetName(j,sheetName);
//                }
                int rowIndex = 1;
                Row row = sheet.createRow(rowIndex);
                boolean spanRow = false;
                List<Integer> indexList = new ArrayList<>();
                int colIndex = 0;
                Row secondRow = null;
                for(int i=0;i<columnTitles.size();i++){
                    if(columnTitles.get(i).indexOf(",") > 0){
                        spanRow = true;
                        String[] titles = columnTitles.get(i).split(",");
                        for(int t = 0;t < titles.length;t++){
                            if(t == 0){
                                Cell cell = row.createCell(colIndex);
                                cell.setCellValue(titles[t]);
                                cell.setCellStyle(normalCellStyle);
                                if (colIndex!=colIndex+titles.length-2){
                                    sheet.addMergedRegion(new CellRangeAddress(rowIndex,rowIndex,colIndex,colIndex+titles.length-2));
                                }
                                if(secondRow == null){
                                    secondRow = sheet.createRow(rowIndex+1);
                                }
                            }else{
                                Cell cell = secondRow.createCell(colIndex);
                                cell.setCellValue(titles[t]);
                                cell.setCellStyle(normalCellStyle);
                                colIndex++;
                            }
                        }
                    }else{
                        indexList.add(colIndex);
                        Cell cell = row.createCell(colIndex++);
                        cell.setCellValue(columnTitles.get(i));
                        cell.setCellStyle(normalCellStyle);
                    }
                }
                if(spanRow){
                    for(Integer index : indexList){
                        sheet.addMergedRegion(new CellRangeAddress(rowIndex,rowIndex+1,index,index));
                    }
                    rowIndex++;
                }
                rowIndex++;
                for(int i=0;i<datas.size();i++){
                    List<String> data = datas.get(i);
                    row = sheet.createRow(rowIndex);
                    colIndex = 0;
                    for(int k=0;k<data.size();k++){
                        row.createCell(colIndex++).setCellValue(data.get(k));
                    }
                    rowIndex++;
                }
            }
            if(os != null){
                workbook.write(os);
                os.flush();
                os.close();
            }else {
                workbook.write(fos);
                fos.flush();
                fos.close();
            }
        }catch (Exception e){
            e.printStackTrace();
        }
    }

    public static void createExportExcelWithHead(Map<String,Object> info){
        try {
            HttpServletResponse response = (HttpServletResponse) info.get("response");
            String fileName = (String)info.get("fileName");
            List<String> columnTitles = (List<String>)info.get("columnTitles");
            List<List<String>> datas = (List<List<String>>) info.get("datas");
            String head = (String) info.get("head");
            List<String> headList = (List<String>)info.get("headList");
            int columnWidth = 16;
            if(info.containsKey("columnWidth")){
                columnWidth = (Integer)info.get("columnWidth");
            }
            List<Integer> columnWidthList = (List<Integer>)info.get("columnWidthList");//各列列宽
            short rowHeight = 400;
            if(info.containsKey("rowHeight")){//行高
                rowHeight = Short.parseShort(info.get("rowHeight").toString());
            }
            OutputStream os = null;
            FileOutputStream fos = null;
            if(response != null){
                os = response.getOutputStream();// 取得输出流
                response.reset();// 清空输出流
                response.setHeader("Content-disposition", "attachment; filename="
                        + new String((fileName + CommonUtil.generateTimestampNo()).getBytes("GB2312"), "8859_1")
                        + ".xls");// 设定输出文件头
                response.setContentType("application/msexcel");// 定义输出类型
            }else {
                fos = (FileOutputStream)info.get("outputStream");
            }
            Workbook workbook = new HSSFWorkbook();
            CellStyle cellStyle = workbook.createCellStyle();
            cellStyle.setWrapText(true);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            Sheet sheet = workbook.createSheet();
            sheet.setDefaultColumnWidth(columnWidth);
            if(columnWidthList != null && columnWidthList.size() == columnTitles.size()){
                for(int i=0;i<columnWidthList.size();i++){
                    sheet.setColumnWidth(i,columnWidthList.get(i)*256);
                }
            }
            int rowIndex = 0;
            Row row = sheet.createRow(rowIndex);
            row.setHeight(rowHeight);
            if(StringUtils.isNotBlank(head)) {
            	Cell cell = row.createCell(0);
            	cell.setCellStyle(cellStyle);
            	cell.setCellValue(fileName);
            	//合并长度并
            	sheet.addMergedRegion(new CellRangeAddress(0,0,0,columnWidth));
            	rowIndex++;
            	row = sheet.createRow(rowIndex);
            }
            if(headList != null){
                int startHead = 0;
                for (String headStr : headList) {
                    Cell cell = row.createCell(startHead);
                    cell.setCellStyle(cellStyle);
                    cell.setCellValue(headStr.split("_")[0]);
                    int colSize = Integer.parseInt(headStr.split("_")[1]);
                    //合并长度并
                    sheet.addMergedRegion(new CellRangeAddress(0,0,startHead,startHead+colSize-1));
                    startHead = startHead + colSize;
                }
                rowIndex++;
                row = sheet.createRow(rowIndex);
            }
            for(int i=0;i<columnTitles.size();i++){
                row.createCell(i).setCellValue(columnTitles.get(i));
            }
            rowIndex++;
            for(int i=0;i<datas.size();i++){
                List<String> data = datas.get(i);
                row = sheet.createRow(rowIndex);
                row.setHeight(rowHeight);
                for(int j=0;j<data.size();j++){
                    Cell cell = row.createCell(j);
                    cell.setCellStyle(cellStyle);
                    cell.setCellValue(data.get(j));
                }
                rowIndex++;
            }
            if(os != null){
                workbook.write(os);
                os.flush();
                os.close();
            }else {
                workbook.write(fos);
                fos.flush();
                fos.close();
            }
        }catch (Exception e){
            e.printStackTrace();
        }
    }

    /**
     * 创建导出excel
     * @param info
     * @return
     */
    public static void createExportExcelToRed(Map<String,Object> info){
        try {
            HttpServletResponse response = (HttpServletResponse) info.get("response");
            String fileName = (String)info.get("fileName");
            List<String> columnTitles = (List<String>)info.get("columnTitles");
            List<List<String>> datas = (List<List<String>>) info.get("datas");
            int columnWidth = 16;
            if(info.containsKey("columnWidth")){
                columnWidth = (Integer)info.get("columnWidth");
            }
            List<Integer> columnWidthList = (List<Integer>)info.get("columnWidthList");//各列列宽
            short rowHeight = 400;
            if(info.containsKey("rowHeight")){//行高
                rowHeight = Short.parseShort(info.get("rowHeight").toString());
            }
            OutputStream os = null;
            FileOutputStream fos = null;
            if(response != null){
                os = response.getOutputStream();// 取得输出流
                response.reset();// 清空输出流
                response.setHeader("Content-disposition", "attachment; filename="
                        + new String((fileName + CommonUtil.generateTimestampNo()).getBytes("GB2312"), "8859_1")
                        + ".xls");// 设定输出文件头
                response.setContentType("application/msexcel");// 定义输出类型
            }else {
                fos = (FileOutputStream)info.get("outputStream");
            }
            Workbook workbook = new HSSFWorkbook();
            CellStyle cellStyle = workbook.createCellStyle();
            cellStyle.setWrapText(true);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            Sheet sheet = workbook.createSheet();
            sheet.setDefaultColumnWidth(columnWidth);
            if(columnWidthList != null && columnWidthList.size() == columnTitles.size()){
                for(int i=0;i<columnWidthList.size();i++){
                    sheet.setColumnWidth(i,columnWidthList.get(i)*256);
                }
            }

            int rowIndex = 0;
            Row row = sheet.createRow(rowIndex);
            row.setHeight(rowHeight);

            Font font = workbook.createFont();
            font.setColor(XSSFFont.COLOR_RED); //红色
            HSSFRichTextString ts;
            for(int i=0;i<columnTitles.size();i++){
                String value = columnTitles.get(i);
                int length = value.length();
                row.createCell(i).setCellValue(value);
                if(i == 1){
                    ts = new HSSFRichTextString(value);
                    ts.applyFont(0,length,font); //0起始索引,5结束索引    标题长度
                    row.getCell(i).setCellValue(ts); //i为第几列，row为第几行
                }else if(i == 2){
                    ts = new HSSFRichTextString(value);
                    ts.applyFont(0,length,font); //0起始索引,3结束索引    标题长度
                    row.getCell(i).setCellValue(ts); //i为第几列，row为第几行
                }
            }
            rowIndex++;
            for(int i=0;i<datas.size();i++){
                List<String> data = datas.get(i);
                row = sheet.createRow(rowIndex);
                row.setHeight(rowHeight);
                for(int j=0;j<data.size();j++){
                    Cell cell = row.createCell(j);
                    cell.setCellStyle(cellStyle);
                    cell.setCellValue(data.get(j));
                }
                rowIndex++;
            }
            if(os != null){
                workbook.write(os);
                os.flush();
                os.close();
            }else {
                workbook.write(fos);
                fos.flush();
                fos.close();
            }
        }catch (Exception e){
            e.printStackTrace();
        }
    }

    public static void createOrgDetailsExportExcel(Map<String,Object> info){
        try {
            HttpServletResponse response = (HttpServletResponse) info.get("response");
            String fileName = (String)info.get("fileName");
            List<Map<String,Object>> dataList = (List<Map<String,Object>>) info.get("dataList");
            int columnWidth = 16;
            if(info.containsKey("columnWidth")){
                columnWidth = (Integer)info.get("columnWidth");
            }
            Boolean split = true;
            if(info.get("noSplit") != null && info.get("noSplit").toString().equals("1")){
                split = false;
            }
            OutputStream os = null;
            FileOutputStream fos = null;
            if(response != null){
                os = response.getOutputStream();// 取得输出流
                response.reset();// 清空输出流
                response.setHeader("Content-disposition", "attachment; filename="
                        + new String((fileName + CommonUtil.generateTimestampNo()).getBytes("GB2312"), "8859_1")
                        + ".xlsx");// 设定输出文件头
                response.setContentType("application/msexcel");// 定义输出类型
            }else {
                fos = (FileOutputStream)info.get("outputStream");
            }
            Workbook workbook = new HSSFWorkbook();
            CellStyle normalCellStyle = workbook.createCellStyle();
            normalCellStyle.setAlignment(HorizontalAlignment.CENTER);
            normalCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

            Map<String,Object> dataMap = null;
            for (int j=0;j<dataList.size();j++){
                dataMap = dataList.get(j);
                List<String> columnTitles = (List<String>)dataMap.get("columnTitles");
                Map<String, List<List<String>>> datas = (Map<String, List<List<String>>>) dataMap.get("datas");
                Integer sizes = (Integer) dataMap.get("sizes");
                String sheetName = (String)dataMap.get("sheetName");
                Sheet sheet = workbook.createSheet();
                Row rowOne = sheet.createRow(0);
                Cell cellOne = rowOne.createCell(0);
                //todo
                sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, sizes));
                cellOne.setCellValue(sheetName);
                cellOne.setCellStyle(normalCellStyle);
                sheet.setDefaultColumnWidth(columnWidth);
                int rowIndex = 1;
                Row row = sheet.createRow(rowIndex);
                boolean spanRow = false;
                List<Integer> indexList = new ArrayList<>();
                int colIndex = 0;
                Row secondRow = null;
                for(int i=0;i<columnTitles.size();i++){
                    if(columnTitles.get(i).indexOf(",") > 0){
                        spanRow = true;
                        String[] titles = columnTitles.get(i).split(",");
                        for(int t = 0;t < titles.length;t++){
                            if(t == 0){
                                Cell cell = row.createCell(colIndex);
                                cell.setCellValue(titles[t]);
                                cell.setCellStyle(normalCellStyle);
                                if (colIndex!=colIndex+titles.length-2){
                                    sheet.addMergedRegion(new CellRangeAddress(rowIndex,rowIndex,colIndex,colIndex+titles.length-2));
                                }
                                if(secondRow == null){
                                    secondRow = sheet.createRow(rowIndex+1);
                                }
                            }else{
                                Cell cell = secondRow.createCell(colIndex);
                                cell.setCellValue(titles[t]);
                                cell.setCellStyle(normalCellStyle);
                                colIndex++;
                            }
                        }
                    }else{
                        indexList.add(colIndex);
                        Cell cell = row.createCell(colIndex++);
                        cell.setCellValue(columnTitles.get(i));
                        cell.setCellStyle(normalCellStyle);
                    }
                }
                if(spanRow){
                    for(Integer index : indexList){
                        sheet.addMergedRegion(new CellRangeAddress(rowIndex,rowIndex+1,index,index));
                    }
                    rowIndex++;
                }
                rowIndex++;
                for(Map.Entry<String, List<List<String>>> entry : datas.entrySet()){
                    Integer firstRow = rowIndex;
                    List<List<String>> value = entry.getValue();
                    for(int i=0;i<value.size();i++){
                        List<String> data = value.get(i);
                        row = sheet.createRow(rowIndex);
                        colIndex = 0;
                        for(int k=0;k<data.size();k++){
                            row.createCell(colIndex++).setCellValue(data.get(k));
                        }
                        rowIndex++;
                    }
                    if (rowIndex-1-firstRow>=1){
                        CellRangeAddress region=new CellRangeAddress(firstRow, rowIndex-1, 0, 0);
                        sheet.addMergedRegion(region);
                    }

                }

            }
            if(os != null){
                workbook.write(os);
                os.flush();
                os.close();
            }else {
                workbook.write(fos);
                fos.flush();
                fos.close();
            }
        }catch (Exception e){
            e.printStackTrace();
        }
    }

}
