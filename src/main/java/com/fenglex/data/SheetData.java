package com.fenglex.data;

import lombok.Data;

import java.util.List;

/**
 * @author haifeng
 * @version 1.0
 * @date 2020/7/24 15:22
 */
@Data
public class SheetData {
    private String name;
    private int index;
    private List<List<CellData>> data;
}
