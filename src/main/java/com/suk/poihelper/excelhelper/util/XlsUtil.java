package com.suk.poihelper.excelhelper.util;


import com.google.common.base.Strings;
import com.suk.poihelper.excelhelper.Interface.OnImportProgressListener;
import com.suk.poihelper.excelhelper.annotation.ExcelFileAttr;
import com.suk.poihelper.excelhelper.annotation.ExcelTypeAttr;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * Author  Administer
 * Time    2016/12/23 16:36
 * Des     ${TODO}
 */

public class XlsUtil {

    public static final String EMAIL_PATTERN =
            "^[_A-Za-z0-9-\\+]+(\\.[_A-Za-z0-9-]+)*@"
                    + "[A-Za-z0-9-]+(\\.[A-Za-z0-9]+)*(\\.[A-Za-z]{2,})$";

    /**
     * 导出excel表
     *
     * @param lists    一个list对应一个excel表
     * @param filePath 导出路径
     */
    public static boolean export(String filePath, List... lists) {
        OutputStream stream = null;
        if (lists == null || lists.length == 0) {
            return false;
        }
        try {
            //xls对象
            HSSFWorkbook workbook = new HSSFWorkbook();
            for (List list : lists) {
                if (list == null || list.size() == 0) {
                    continue;
                }
                //sheet
                HSSFSheet sheet = workbook.createSheet();
                //获取第一个item
                Object obj1 = list.get(0);
                Class<?> clazz = obj1.getClass();// 获取集合中的对象类型
                String title;
                HSSFCellStyle titleStyle = workbook.createCellStyle();
                HSSFFont titleFont = workbook.createFont();
                if (clazz.isAnnotationPresent(ExcelTypeAttr.class)) { //获取表的标题
                    ExcelTypeAttr clazzAnnotation = clazz.getAnnotation(ExcelTypeAttr.class);
                    //title
                    title = getTitleByAnno(clazzAnnotation);
                    //style
                    short heightInPoints = clazzAnnotation.fontHeightInPoints();
                    short alignment = clazzAnnotation.alignment();
                    short verticalAlignment = clazzAnnotation.verticalAlignment();
                    titleFont.setFontHeightInPoints(heightInPoints);
                    titleStyle.setFont(titleFont);
                    titleStyle.setAlignment(alignment);
                    titleStyle.setVerticalAlignment(verticalAlignment);
                } else {
                    title = clazz.getName();
                }
                Field[] fds = clazz.getDeclaredFields();// 获取他的字段数组
                System.out.print("ClassName:" + clazz.getName());
                for (Field f : fds) {
                    System.out.print("fdName:" + f.getName());
                }
                //标题行
                HSSFRow row0 = sheet.createRow(0);
                HSSFCell titleCell = row0.createCell(0);
                sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, fds.length - 1));
                titleCell.setCellValue(title);
                titleCell.setCellStyle(titleStyle);
                //属性行
                HSSFRow row1 = sheet.createRow(1);
                //字段排序
                ArrayList<Field> fdList = new ArrayList<>(Arrays.asList(fds));
                Map<Integer, Field> index = new HashMap<>();
                for (Iterator<Field> it = fdList.iterator(); it.hasNext(); ) {
                    Field fd = it.next();
                    if (fd.isAnnotationPresent(ExcelFileAttr.class)) {
                        ExcelFileAttr annotation = fd.getAnnotation(ExcelFileAttr.class);
                        int column = annotation.column();
                        if (column != -1) {
                            // TODO: 2017/5/15 等于-1的情况
                            //                            it.remove();
                            index.put(column, fd);
                        }
                    }
                }
                for (Map.Entry<Integer, Field> entrySet : index.entrySet()) {
                    fdList.set(entrySet.getKey(), entrySet.getValue());
                }
                //设置中文名,样式
                for (int i = 0; i < fdList.size(); i++) {
                    Field fd = fdList.get(i);
                    HSSFCell cell = row1.createCell(i);
                    if (fd.isAnnotationPresent(ExcelFileAttr.class)) {
                        ExcelFileAttr annotation = fd.getAnnotation(ExcelFileAttr.class);
                        //属性值
                        cell.setCellValue(getNameByAnno(annotation));
                        //style
                        HSSFCellStyle style = workbook.createCellStyle();
                        HSSFFont font = workbook.createFont();
                        short heightInPoints = annotation.fieldFontHeightInPoints();
                        short alignment = annotation.alignment();
                        short verticalAlignment = annotation.verticalAlignment();
                        short color = annotation.fieldFontColor();
                        font.setColor(color);
                        font.setFontHeightInPoints(heightInPoints);
                        style.setFont(font);
                        style.setAlignment(alignment);
                        style.setVerticalAlignment(verticalAlignment);
                        cell.setCellStyle(style);
                    } else {
                        cell.setCellValue(fd.getName());
                    }
                }
                //内容,从第三行开始
                int rowIndex = 2;
                HashMap<Field, HSSFCellStyle> styleMap = new HashMap<>();
                for (Field field : fdList) {
                    if (field.isAnnotationPresent(ExcelFileAttr.class)) {
                        ExcelFileAttr annotation = field.getAnnotation(ExcelFileAttr.class);
                        short heightInPoints = annotation.valueFontHeightInPoints();
                        short alignment = annotation.alignment();
                        short verticalAlignment = annotation.verticalAlignment();
                        HSSFCellStyle style = workbook.createCellStyle();
                        HSSFFont font = workbook.createFont();
                        font.setFontHeightInPoints(heightInPoints);
                        style.setFont(font);
                        style.setAlignment(alignment);
                        style.setVerticalAlignment(verticalAlignment);
                        styleMap.put(field, style);
                    }
                }
                for (Object obj : list) {
                    HSSFRow row = sheet.createRow(rowIndex++); //
                    for (int i = 0; i < fdList.size(); i++) {// 遍历该数组
                        HSSFCell cell = row.createCell(i);
                        Field field = fdList.get(i);
                        System.out.print("name:  " + upperCase(field.getName()));// 得到属性名
                        Method method = clazz.getMethod("get" + upperCase(field.getName()));
                        if (method != null) {
                            Object value = method.invoke(obj);
                            System.out.print("value:  " + value);// 得到属性值
                            //属性值
                            cell.setCellValue(value == null ? "" : value.toString());
                            //style
                            HSSFCellStyle style = styleMap.get(field);
                            if (style != null) {
                                cell.setCellStyle(style);
                            }
                        }
                    }
                }
            }
            System.out.print("filePath:" + filePath);
            File file = new File(filePath);
            //写入
            stream = new FileOutputStream(file);
            workbook.write(stream);
            stream.close();
            //返回
            return true;
        } catch (Exception e) { //NoSuchMethod,InvocationTarget,IllegalAccess,FileNotFound,IO
            e.printStackTrace();
        }
        if (stream != null) {
            try {
                stream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return false;
    }

    /**
     * @param filePath 文档路径
     * @param clazzs   每个sheet对应一个Class
     * @return key = 数据 , value = 错误信息 (n行n列,错误信息提示)
     */
    public static HashMap<List<Object>, HashMap<Map<Integer, Integer>, String>> improt(String filePath, OnImportProgressListener onProgressListener, Class<?>... clazzs) {
        //        ArrayList<List<Object>> itemsList = new ArrayList<>();
        HashMap<List<Object>, HashMap<Map<Integer, Integer>, String>> resultMaps = new HashMap<>();
        try {
            InputStream is = new FileInputStream(filePath);
            HSSFWorkbook book = new HSSFWorkbook(is);
            if (book.getNumberOfSheets() != clazzs.length) {
                return null;
            }
            for (Sheet sheet : book) { //如果全部sheet行数小于3,则为空不能导入
                if (sheet.getLastRowNum() > 1) {
                    continue;
                }
                return null;
            }

            for (Sheet sheet : book) { //遍历sheet

                //                long l = System.currentTimeMillis();
                //sheel的索引
                int sheetIndex = book.getSheetIndex(sheet);
                //有效的行数
                int count = getvalidRowCount(sheet);
                //                Log.d("520", "耗时:" + (System.currentTimeMillis() - l));


                String sheetTitle = sheet.getRow(0).getCell(0).getStringCellValue();
                System.out.print("title:" + sheetTitle);
                for (Class<?> clazz : clazzs) { //遍历找到符合当前sheet标题的JavaBean
                    String title = null;
                    if (clazz.isAnnotationPresent(ExcelTypeAttr.class) && clazzs.length > 1) {
                        title = trim(getTitleByAnno(clazz.getAnnotation(ExcelTypeAttr.class)));
                    } else if (clazzs.length > 1) { //只有clazzs大于1才需要表头区分
                        title = clazz.getName();
                    } else if (clazzs.length == 1) { //若clazzs只有一个且没有注解,那不用区分直接导入
                        title = sheetTitle;
                    }
                    if (sheetTitle.equals(title)) {
                        //字段数组根据第二行排序
                        ArrayList<Field> feds = new ArrayList<>(Arrays.asList(clazz.getDeclaredFields()));
                        HashMap<Integer, Field> map = new HashMap<>();
                        Row row1 = sheet.getRow(1);
                        for (int z = 0; z < feds.size(); z++) {
                            Cell cell = row1.getCell(z);
                            if (cell == null) {
                                map.put(z, null);
                                continue;
                            }
                            String cellStr = cell.getStringCellValue();
                            System.out.print("cellStr:" + cellStr);
                            for (Field fed : feds) {
                                String fedName;
                                if (fed.isAnnotationPresent(ExcelFileAttr.class)) {
                                    ExcelFileAttr anno = fed.getAnnotation(ExcelFileAttr.class);
                                    fedName = getNameByAnno(anno);
                                    if (cellStr.equals(fedName)) {
                                        map.put(z, fed);
                                        break; // TODO: 2017/3/16
                                    }
                                } else {
                                    fedName = fed.getName();
                                    if (cellStr.equals(fedName)) {
                                        map.put(z, fed);
                                        break; // TODO: 2017/3/16
                                    }
                                }
                            }
                            if (!map.containsKey(z)) {
                                return null;// TODO: 2017/3/16  应付需求,通过字段完全匹配这个规则判断calzz是否匹配excel表
                            }
                        }
                        for (Map.Entry<Integer, Field> entry : map.entrySet()) {
                            feds.set(entry.getKey(), entry.getValue());
                        }
                        System.out.print("field:" + feds);
                        //从第三行开始遍历
                        ArrayList<Object> items = new ArrayList<>();
                        HashMap<Map<Integer, Integer>, String> errorMap = new LinkedHashMap<>();
                        for (int i = 2; i <= sheet.getLastRowNum(); i++) {
                            Row row = sheet.getRow(i);
                            if (isEmptyRow(row)) {
                                continue;
                            }
                            Object item = clazz.newInstance();
                            sec:
                            for (int j = 0; j < feds.size(); j++) {
                                Cell cell = row.getCell(j);
                                String val;
                                if (cell != null) {
                                    //有时候日期会读取成一个数字字符串,不是原字符串
                                    if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC && HSSFDateUtil.isCellDateFormatted(cell)) {
                                        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd", Locale.getDefault());
                                        Date date = HSSFDateUtil.getJavaDate(cell.getNumericCellValue());
                                        val = sdf.format(date);
                                    } else {
                                        cell.setCellType(CellType.STRING);
                                        String stringVal = trim(cell.getStringCellValue());
                                        val = Strings.isNullOrEmpty(stringVal) ? null : stringVal;
                                    }
                                    System.out.print("value:" + val);
                                } else {
                                    val = null;
                                }
                                Field field = feds.get(j);
                                if (field == null) {
                                    continue;
                                } else if (field.isAnnotationPresent(ExcelFileAttr.class)) { //如果某必要字段值缺失,则这行没用了,但要进行错误记录所以item为null,但要继续轮循下去
                                    ExcelFileAttr anno = field.getAnnotation(ExcelFileAttr.class);
                                    String name = getNameByAnno(anno);
                                    if (anno.notNull() && val == null) {
                                        HashMap<Integer, Integer> locationMap = new HashMap<>();
                                        locationMap.put(i + 1, j + 1);
                                        errorMap.put(locationMap, name + ":不能为空");
                                        item = null;
                                        continue;
                                    } else if (val != null) {
                                        if (anno.isWholeNum() && !isNumeric(val)) {
                                            HashMap<Integer, Integer> locationMap = new HashMap<>();
                                            locationMap.put(i + 1, j + 1);
                                            errorMap.put(locationMap, name + ":必须是纯数字");
                                            item = null;
                                            continue;
                                        }
                                        if (anno.maxStrLen() != -1 && getStrlength(val) > anno.maxStrLen()) {
                                            HashMap<Integer, Integer> locationMap = new HashMap<>();
                                            locationMap.put(i + 1, j + 1);
                                            if (isNumeric(val)) {
                                                errorMap.put(locationMap, name + ":长度请限制在" + anno.maxStrLen() + "位数字以内");
                                            } else {
                                                errorMap.put(locationMap, name + ":长度请限制在" + anno.maxStrLen() + "个字符以内");
                                            }
                                            item = null;
                                            continue;
                                        }
                                        if (anno.isEmail() && !isEamil(val)) {
                                            HashMap<Integer, Integer> locationMap = new HashMap<>();
                                            locationMap.put(i + 1, j + 1);
                                            errorMap.put(locationMap, name + ":邮箱格式错误");
                                            item = null;
                                            continue;
                                        }
                                    }
                                }
                                for (Map<Integer, Integer> key : errorMap.keySet()) {
                                    if (key.containsKey(i + 1)) {
                                        continue sec;
                                    }
                                }
                                Class<?> type = field.getType();
                                System.out.print("method:" + "set" + upperCase(field.getName()) + "   " + clazz.getName());
                                Method method = clazz.getMethod("set" + upperCase(field.getName()), type);
                                if (method != null) { //有错不给执行
                                    //todo 暂时只支持接收item所有字段为String类型
                                    method.invoke(item, val);
                                }
                            }
                            if (item != null) {
                                items.add(item);
                                if (onProgressListener != null) {
                                    onProgressListener.onProgress(sheetIndex, count - 2, i - 2, item);
                                }
                            }
                        }
                        resultMaps.put(items, errorMap);
                        //                        itemsList.add(items);
                    }
                }
            }
            for (Map.Entry<List<Object>, HashMap<Map<Integer, Integer>, String>> resultMap : resultMaps.entrySet()) {
                HashMap<Map<Integer, Integer>, String> errorMap = resultMap.getValue();
                for (Map.Entry<Map<Integer, Integer>, String> error : errorMap.entrySet()) {
                    Map<Integer, Integer> location = error.getKey();
                    for (Map.Entry<Integer, Integer> entry : location.entrySet()) {
                        System.out.print("error:" + entry.getKey() + "行" + entry.getValue() + "列  " + "错误信息:" + error.getValue());
                    }
                }
                List<Object> items = resultMap.getKey();
                System.out.print("items:" + items.toString());
            }
            return resultMaps;
        } catch (Exception e) { //NoSuchMethod,InvocationTarget,IllegalAccess,FileNotFound,IO,InstantiationException
            e.printStackTrace();
        }
        return null;
    }

    /**
     * 判断是否为空行
     */
    private static boolean isEmptyRow(Row row) {
        for (Cell cell : row) {
            int cellType = cell.getCellType();
            cell.setCellType(CellType.STRING);
            String value = cell.getStringCellValue();
            cell.setCellType(cellType);
            if (!Strings.isNullOrEmpty(value)) {
                return false;
            }
        }
        return true;
    }

    private static int getvalidRowCount(Sheet sheet) {
        int i = 0;
        for (Row row : sheet) {
            if (!isEmptyRow(row)) {
                i++;
            }
        }
        return i;
    }

    /**
     * 导出一份只有表头和属性名的表
     */
    public static boolean exportEmptyExcel(String filePath, Class clazz) {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFCellStyle titleStyle = workbook.createCellStyle();
        HSSFFont titleFont = workbook.createFont();
        String title;
        //sheet
        HSSFSheet sheet = workbook.createSheet();
        if (clazz.isAnnotationPresent(ExcelTypeAttr.class)) { //获取表的标题
            ExcelTypeAttr clazzAnnotation = (ExcelTypeAttr) clazz.getAnnotation(ExcelTypeAttr.class);
            //title
            title = getTitleByAnno(clazzAnnotation);
            //style
            short heightInPoints = clazzAnnotation.fontHeightInPoints();
            short alignment = clazzAnnotation.alignment();
            short verticalAlignment = clazzAnnotation.verticalAlignment();
            titleFont.setFontHeightInPoints(heightInPoints);
            titleStyle.setFont(titleFont);
            titleStyle.setAlignment(alignment);
            titleStyle.setVerticalAlignment(verticalAlignment);
        } else {
            title = clazz.getName();
        }
        Field[] fds = clazz.getDeclaredFields();// 获取他的字段数组
        //标题行
        HSSFRow row0 = sheet.createRow(0);
        HSSFCell titleCell = row0.createCell(0);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, fds.length - 1));
        titleCell.setCellValue(title);
        titleCell.setCellStyle(titleStyle);
        //属性行
        HSSFRow row1 = sheet.createRow(1);
        //字段排序
        ArrayList<Field> fdList = new ArrayList<>(Arrays.asList(fds));
        Map<Integer, Field> index = new HashMap<>();
        for (Iterator<Field> it = fdList.iterator(); it.hasNext(); ) {
            Field fd = it.next();
            if (fd.isAnnotationPresent(ExcelFileAttr.class)) {
                ExcelFileAttr annotation = fd.getAnnotation(ExcelFileAttr.class);
                int column = annotation.column();
                if (column != -1) {
                    //                    it.remove();
                    // TODO: 2017/5/16
                    index.put(column, fd);
                }
            }
        }
        for (Map.Entry<Integer, Field> entrySet : index.entrySet()) {
            fdList.set(entrySet.getKey(), entrySet.getValue());
        }
        //设置中文名,样式
        for (int i = 0; i < fdList.size(); i++) {
            Field fd = fdList.get(i);
            HSSFCell cell = row1.createCell(i);
            if (fd.isAnnotationPresent(ExcelFileAttr.class)) {
                ExcelFileAttr annotation = fd.getAnnotation(ExcelFileAttr.class);
                //属性值
                cell.setCellValue(getNameByAnno(annotation));
                //style
                HSSFCellStyle style = workbook.createCellStyle();
                HSSFFont font = workbook.createFont();
                short heightInPoints = annotation.fieldFontHeightInPoints();
                short alignment = annotation.alignment();
                short verticalAlignment = annotation.verticalAlignment();
                short color = annotation.fieldFontColor();
                font.setColor(color);
                font.setFontHeightInPoints(heightInPoints);
                style.setFont(font);
                style.setAlignment(alignment);
                style.setVerticalAlignment(verticalAlignment);
                cell.setCellStyle(style);
            } else {
                cell.setCellValue(fd.getName());
            }
        }
        System.out.print("filePath:" + filePath);
        File file = new File(filePath);
        //写入
        try {
            OutputStream stream = new FileOutputStream(file);
            workbook.write(stream);
            stream.close();
            return true;
        } catch (IOException e) {
            e.printStackTrace();
            return false;
        }
    }


    /**
     * 去除前后空格
     *
     * @param str
     * @return
     */
    private static String trim(String str) {
        str = str.trim();
        while (str.startsWith("　")) {
            str = str.substring(1, str.length()).trim();
        }
        while (str.endsWith("　")) {
            str = str.substring(0, str.length() - 1).trim();
        }
        return str;
    }

    private static String getNameByAnno(ExcelFileAttr anno) {
        return anno.nameStr();
    }

    private static String getTitleByAnno(ExcelTypeAttr anno) {
        return anno.titleStr();
    }

    /**
     * @param src 源字符串
     * @return 字符串，将src的第一个字母转换为大写，src为空时返回null
     */
    private static String upperCase(String src) {
        if (src != null) {
            StringBuffer sb = new StringBuffer(src);
            sb.setCharAt(0, Character.toUpperCase(sb.charAt(0)));
            return sb.toString();
        } else {
            return null;
        }
    }

    private static int getStrlength(String value) {
        if (Strings.isNullOrEmpty(value)) {
            return 0;
        }
        int valueLength = 0;
        String chinese = "[\u0391-\uFFE5]";
        /* 获取字段值的长度，如果含中文字符，则每个中文字符长度为2，否则为1 */
        for (int i = 0; i < value.length(); i++) {
            /* 获取一个字符 */
            String temp = value.substring(i, i + 1);
            /* 判断是否为中文字符 */
            if (temp.matches(chinese)) {
                /* 中文字符长度为2 */
                valueLength += 2;
            } else {
                /* 其他字符长度为1 */
                valueLength += 1;
            }
        }
        return valueLength;
    }

    private static boolean isNumeric(String str) {
        if (Strings.isNullOrEmpty(str)) {
            return false;
        }
        for (int i = 0; i < str.length(); i++) {
            if (!Character.isDigit(str.charAt(i))) {
                return false;
            }
        }
        return true;
    }

    private static boolean isEamil(String str) {
        if (Strings.isNullOrEmpty(str)) {
            return false;
        }
        return str.matches(EMAIL_PATTERN);
    }

}
