package com.ywluv.easy_excel;


import com.ywluv.easy_excel.enums.EasyCellStyleEnum;

import java.util.*;

/**
 * packageName    : com.ywluv.easyExcel
 * fileName       : EasyRow.java
 * author         : MYH
 * date           : 2025-08-25
 * description    :
 * ===========================================================
 * DATE              AUTHOR             NOTE
 * -----------------------------------------------------------
 * 2025-08-25        MYH       최초 생성
 */
public class EasyRow {

    private final List<Cell> cellValues;
    private Set<EasyCellStyleEnum> rowStyles;

    private EasyRow(Set<EasyCellStyleEnum> rowStyles, List<Cell> cellValues) {
        this.rowStyles = rowStyles;
        this.cellValues = Collections.unmodifiableList(cellValues);
    }

    public static Builder builder() {
        return new Builder();
    }

    public List<Cell> getCells() {
        return cellValues;
    }

    public Set<EasyCellStyleEnum> getRowStyles() {
        return rowStyles;
    }

    public static class Builder {
        private Set<EasyCellStyleEnum> tempRowStyles = new HashSet<>();
        private List<Cell> tempCells = new ArrayList<>();

        public Builder addCell(Set<EasyCellStyleEnum> cellStyles, Object value) {
            String strValue;

            if (value == null) {
                strValue = "";
            } else if (value instanceof Boolean) {
                strValue = (Boolean) value ? "Y" : "N";
            } else {
                strValue = String.valueOf(value);
            }

            tempCells.add(new Cell(cellStyles, strValue));
            return this;
        }

        public Builder addCells(Set<EasyCellStyleEnum> cellStyles, Object... values) {
            for (Object value : values) {
                String strValue;
                if (value == null) {
                    strValue = "";
                } else if (value instanceof Boolean) {
                    strValue = ((Boolean) value) ? "Y" : "N";
                } else {
                    strValue = String.valueOf(value);
                }

                tempCells.add(new Cell(cellStyles, strValue));
            }

            return this;
        }

        public Builder rowStyle(Set<EasyCellStyleEnum> rowStyles){
            this.tempRowStyles = rowStyles;
            return this;
        }

        public Builder addRowStyle(EasyCellStyleEnum rowStyle) {
            if (rowStyle != null) {
                this.tempRowStyles.add(rowStyle);
            }
            return this;
        }

        public Builder addRowStyles(Set<EasyCellStyleEnum> rowStyles) {
            if (rowStyles != null && !rowStyles.isEmpty()) {
                this.tempRowStyles.addAll(rowStyles);
            }
            return this;
        }

        public EasyRow build() {
            return new EasyRow(tempRowStyles, new ArrayList<>(tempCells));
        }
    }

    public static class Cell {
        private final Set<EasyCellStyleEnum> cellStyles;
        private final String value;

        public Cell(Set<EasyCellStyleEnum> cellStyles, String value) {
            this.cellStyles = cellStyles;
            this.value = value;
        }

        public String getValue() {
            return value;
        }

        public Set<EasyCellStyleEnum> getCellStyles() {
            return cellStyles;
        }
    }
}

