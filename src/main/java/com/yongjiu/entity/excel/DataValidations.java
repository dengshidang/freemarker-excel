package com.yongjiu.entity.excel;

/**
 * @author dengsd
 * @date 2023/1/13 11:00
 */

public class DataValidations {
    private String valueText;
    private int startRow;
    private int endRow;
    private int startCol;
    private int endCol;

    public DataValidations(String valueText, int startRow, int endRow, int startCol, int endCol) {
        this.valueText = valueText;
        this.startRow = startRow;
        this.endRow = endRow;
        this.startCol = startCol;
        this.endCol = endCol;
    }

    public String getValueText() {
        return valueText;
    }

    public void setValueText(String valueText) {
        this.valueText = valueText;
    }

    public int getStartRow() {
        return startRow;
    }

    public void setStartRow(int startRow) {
        this.startRow = startRow;
    }

    public int getEndRow() {
        return endRow;
    }

    public void setEndRow(int endRow) {
        this.endRow = endRow;
    }

    public int getStartCol() {
        return startCol;
    }

    public void setStartCol(int startCol) {
        this.startCol = startCol;
    }

    public int getEndCol() {
        return endCol;
    }

    public void setEndCol(int endCol) {
        this.endCol = endCol;
    }
}
