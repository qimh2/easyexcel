package com.alibaba.easyexcel.test.demo.web;

import com.alibaba.excel.annotation.write.style.ColumnWidth;
import com.alibaba.excel.annotation.write.style.ContentRowHeight;
import com.alibaba.excel.annotation.write.style.HeadRowHeight;
import java.util.Date;

import com.alibaba.excel.annotation.ExcelProperty;

import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.Getter;
import lombok.Setter;

/**
 * 基础数据类
 *
 * @author Jiaju Zhuang
 **/
@Data
@ContentRowHeight(50) // 设置 Cell 高度 为50
@HeadRowHeight(40)    //  设置表头 高度 为 40
public class DownloadData {
    @ExcelProperty(value = {"字符串标题"},index = 2)
    @ColumnWidth(20)
    private String string;

    @ExcelProperty(value = "日期标题",index = 1)
    @ColumnWidth(20)
    private Date date;

    @ExcelProperty(value = "数字标题" ,index = 0)
    @ColumnWidth(20)
    private Double doubleData;
}
