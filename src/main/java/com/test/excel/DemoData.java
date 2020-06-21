package com.test.excel;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.HeadStyle;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@AllArgsConstructor
@NoArgsConstructor
@Builder
@HeadStyle(fillForegroundColor = 9)
public class DemoData {
    @ExcelProperty({"标题", "部位", "项目"})
    private String xm;
    @ExcelProperty({"标题", "气象站", "平均速度"})
    private String pjsd;
    @ExcelProperty({"标题", "气象站", "空气温度"})
    private String kqwd;
    @ExcelProperty({"标题", "轮毂", "轮毂转速"})
    private String lgzs;
    @ExcelProperty({"标题", "轮毂", "叶片角度", "叶片1"})
    private String ypjd1;
    @ExcelProperty({"标题", "轮毂", "叶片角度", "叶片2"})
    private String ypjd2;
    @ExcelProperty({"标题", "轮毂", "叶片角度", "叶片3"})
    private String ypjd3;
    @ExcelProperty({"标题", "轮毂", "液缸压力", "叶片1"})
    private String ypyg1;
    @ExcelProperty({"标题", "轮毂", "液缸压力", "叶片2"})
    private String ypyg2;
    @ExcelProperty({"标题", "轮毂", "液缸压力", "叶片3"})
    private String ypyg3;
    @ExcelProperty({"标题", "轮毂", "泵蓄能器压力"})
    private String bxnqyl;
}
