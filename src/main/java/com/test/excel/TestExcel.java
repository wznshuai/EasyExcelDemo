package com.test.excel;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import com.alibaba.excel.write.style.HorizontalCellStyleStrategy;
import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.google.common.collect.Lists;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.text.SimpleDateFormat;
import java.util.*;

public class TestExcel {
    public static void main(String[] args) {
        testDynamicExcel();
    }

    private static void testDynamicExcel(){
        String excelName = System.getProperty("user.dir") + "/test2.xls";
        JSONObject object = JSON.parseObject(Constats.json);
        // 头的策略
        WriteCellStyle headWriteCellStyle = new WriteCellStyle();
        // 背景设置为红色
        headWriteCellStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
//        WriteFont headWriteFont = new WriteFont();
//        headWriteFont.setFontHeightInPoints((short)20);
//        headWriteCellStyle.setWriteFont(headWriteFont);
        // 内容的策略
        WriteCellStyle contentWriteCellStyle = new WriteCellStyle();
        // 这里需要指定 FillPatternType 为FillPatternType.SOLID_FOREGROUND 不然无法显示背景颜色.头默认了 FillPatternType所以可以不指定
        contentWriteCellStyle.setFillPatternType(FillPatternType.SOLID_FOREGROUND);
        // 背景绿色
        contentWriteCellStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
        WriteFont contentWriteFont = new WriteFont();
        // 字体大小
        contentWriteFont.setFontHeightInPoints((short)20);
        contentWriteCellStyle.setWriteFont(contentWriteFont);
        // 这个策略是 头是头的样式 内容是内容的样式 其他的策略可以自己实现
        HorizontalCellStyleStrategy horizontalCellStyleStrategy =
                new HorizontalCellStyleStrategy(headWriteCellStyle, contentWriteCellStyle);

        EasyExcel.write(excelName).head(addHeader(object)).registerWriteHandler(horizontalCellStyleStrategy).sheet().doWrite(Lists.newArrayList());
    }

    private static void testFixExcel(){
        String excelName = System.getProperty("user.dir") + "/test.xls";
        String dateTime = new SimpleDateFormat("yyyy年MM月dd日 HH时").format(new Date());
        Set<String> includeColumnFiledNames = new HashSet<String>();
        includeColumnFiledNames.add("项目");
        EasyExcel.write(excelName).head(DemoData.class).sheet("表格1")
                .excludeColumnFiledNames(includeColumnFiledNames)
                .registerWriteHandler(new CustomerHandler("xxx电场参数表    时间: " + dateTime)).doWrite(listData());
    }



    private static List<DemoData> listData(){
        List<DemoData> demoDataList = new ArrayList<>();
        demoDataList.add(DemoData.builder().xm("定值").bxnqyl("1").kqwd("2").lgzs("3").build());
        demoDataList.add(DemoData.builder().xm("风机1").bxnqyl("1").kqwd("2").lgzs("3").build());
        demoDataList.add(DemoData.builder().xm("风机2").bxnqyl("1").kqwd("2").lgzs("3").build());
        return demoDataList;
    }

    private static List<List<String>> addHeader(JSONObject jsonObject){
        List<List<String>> headers = new ArrayList<>();
        List<String> column = new ArrayList<>();
        column.add(jsonObject.getString("name"));
        JSONArray children = jsonObject.getJSONArray("children");
        if(null != children && children.size() > 0){
            for (Object obj : children) {
                eachJson(headers, column, (JSONObject) obj);
            }
        }
        return headers;
    }

    private static void eachJson(List<List<String>> headers, List<String> column, JSONObject jsonObject){
        List<String> newColumn = new ArrayList<>(column);
        newColumn.add(jsonObject.getString("name"));

        JSONArray children = jsonObject.getJSONArray("children");
        if(null != children && children.size() > 0){
            for (Object obj : children) {
                 eachJson(headers, newColumn, (JSONObject) obj);
            }
        }else {
            headers.add(newColumn);
        }
    }
}
