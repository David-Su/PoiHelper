package com.suk.poihelper.excelhelper.annotation;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Author  Administer
 * Time    2017/11/24 17:26
 * Des
 */

@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD})
public @interface ExcelFileAttr {

    /**
     * 导出到excel表的属性中文名的字符串
     */
    String nameStr() default "";

    /**
     * 配置列的名称,对应A,B,C,D....
     */
    int column() default -1;

    /**
     * 是否可以为空
     */
    boolean notNull() default false;

    /**
     * 字段字体大小
     */
    short fieldFontHeightInPoints() default (short) 13;

    /**
     * 字段字体颜色
     */
    short fieldFontColor() default HSSFFont.BOLDWEIGHT_BOLD;

    /**
     * 字段值体大小
     */
    short valueFontHeightInPoints() default (short) 10;

    /**
     * 水平对齐方式
     */
    short alignment() default HSSFCellStyle.ALIGN_CENTER_SELECTION;

    /**
     * 竖直对齐方式
     */
    short verticalAlignment() default HSSFCellStyle.VERTICAL_CENTER;

    /**
     * 规定的最大字符数
     */
    int maxStrLen() default -1;

    /**
     * 是否全部为数字
     */
    boolean isWholeNum() default false;

    /**
     * 是否要求为邮箱格式
     */
    boolean isEmail() default false;
}