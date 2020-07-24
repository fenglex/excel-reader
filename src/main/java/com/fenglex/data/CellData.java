package com.fenglex.data;

import lombok.Data;
import org.apache.poi.ss.util.CellReference;

/**
 * @author haifeng
 * @version 1.0
 * @date 2020/7/24 15:23
 */
@Data
public class CellData {
    /**
     * 行号（从1开始）
     */
    private int row;
    /**
     * 列号（从1开始）
     */
    private int column;
    /**
     * 获取A1格式excel信息
     */
    private String cellReference;
    private String value;

    public CellData(int row, int column, String value) {
        this.row = row;
        this.column = column;
        this.value = value;
    }

    public CellData(int row, int column) {
        this.row = row;
        this.column = column;
        this.value = "";
    }

    public String getCellReference() {
        return new CellReference(this.row, this.column).formatAsString();
    }
}
