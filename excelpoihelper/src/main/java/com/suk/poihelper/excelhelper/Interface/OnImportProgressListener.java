package com.suk.poihelper.excelhelper.Interface;

/**
 * Author  Administer
 * Time    2017/12/7 19:54
 * Des
 */

public interface OnImportProgressListener {

    /**
     *
     * @param sheetIndex 分页索引
     * @param rowCount 总行数(不算表头和属性行)
     * @param rowIndex 当前读到的行索引(不算表头和属性行)
     * @param item 当前行封装成的Javabean
     */
    void onProgress(int sheetIndex, int rowCount, int rowIndex, Object item);
}
