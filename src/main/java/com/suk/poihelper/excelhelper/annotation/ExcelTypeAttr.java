package com.suk.poihelper.excelhelper.annotation;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Author  Administer
 * Time    2017/11/24 17:27
 * Des
 */

@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.TYPE})
public @interface ExcelTypeAttr {

    /**
     * 导出到excel表的第一行标题中文字符串
     */
    String titleStr() default "";

    /**
     * 字体大小
     */
    short fontHeightInPoints() default (short) 15;

    /**
     * 水平对齐方式
     */
    short alignment() default HSSFCellStyle.ALIGN_CENTER_SELECTION;

    /**
     * 竖直对齐方式
     */
    short verticalAlignment() default HSSFCellStyle.VERTICAL_CENTER;

}