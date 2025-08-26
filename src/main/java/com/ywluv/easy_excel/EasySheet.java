package com.ywluv.easy_excel;


import com.ywluv.easy_excel.enums.EasyCellStyleEnum;

import java.util.*;

public class EasySheet {

    private final String sheetName;
    private final Set<EasyCellStyleEnum> sheetStyles;
    private final List<EasyRow> easyRows;
    private final Integer freezeCol;
    private final Integer freezeRow;

    private EasySheet(String sheetName, Set<EasyCellStyleEnum> sheetStyles, List<EasyRow> easyRows, Integer freezeCol , Integer freezeRow) {
        this.sheetName = sheetName;
        this.sheetStyles = sheetStyles;
        this.easyRows = Collections.unmodifiableList(easyRows != null ? new ArrayList<>(easyRows) : new ArrayList<>());
        this.freezeCol = freezeCol;
        this.freezeRow = freezeRow;
    }

    public static Builder builder() {
        return new Builder();
    }

    public String getSheetName() {
        if (sheetName == null) {
            return null;
        }
        String clean = sheetName.replaceAll("[\\\\/?*\\[\\]:]", "_");
        return clean.length() > 31 ? clean.substring(0, 31) : clean;
    }


    public Set<EasyCellStyleEnum> getSheetStyles() {
        return sheetStyles;
    }

    public List<EasyRow> getRows() {
        return easyRows;
    }

    public Integer getFreezeCol(){ return freezeCol;}


    public Integer getFreezeRow(){ return freezeRow;}

    public static class Builder {
        private String sheetName;
        private Set<EasyCellStyleEnum> tempSheetStyles = new HashSet<>();
        private List<EasyRow> tempEasyRows = new ArrayList<>();
        private int freezeCol = 0;
        private int freezeRow = 0;

        public Builder sheetName(String sheetName) {
            this.sheetName = sheetName;
            return this;
        }

        public Builder sheetStyle(Set<EasyCellStyleEnum> sheetStyles){
            this.tempSheetStyles = sheetStyles;
            return this;
        }

        public Builder addSheetStyle(EasyCellStyleEnum sheetStyle) {
            if (sheetStyle != null) {
                this.tempSheetStyles.add(sheetStyle);
            }
            return this;
        }

        public Builder addSheetStyles(Set<EasyCellStyleEnum> sheetStyles) {
            if (sheetStyles != null && !sheetStyles.isEmpty()) {
                this.tempSheetStyles.addAll(sheetStyles);
            }
            return this;
        }

        public Builder addRow(EasyRow easyRow) {
            if (easyRow != null) {
                tempEasyRows.add(easyRow);
            }
            return this;
        }

        public Builder addRow(EasyRow easyRow, Integer index) {
            if (easyRow != null) {
                int pos;
                if (index == null || index < 0) {
                    pos = 0;
                } else if (index > tempEasyRows.size()) {
                    pos = tempEasyRows.size();
                } else {
                    pos = index;
                }
                tempEasyRows.add(pos, easyRow);
            }
            return this;
        }

        public Builder addRowsAtEnd(List<EasyRow> easyRows) {
            if (easyRows != null && !easyRows.isEmpty()) {
                tempEasyRows.addAll(easyRows);
            }
            return this;
        }

        public Builder addRowsAtStart(List<EasyRow> easyRows) {
            if (easyRows != null && !easyRows.isEmpty()) {
                tempEasyRows.addAll(0, easyRows);
            }
            return this;
        }

        public Builder freezePane(int freezeCol, int freezeRow) {
            if (freezeCol >= 0) this.freezeCol = freezeCol;
            if (freezeRow >= 0) this.freezeRow = freezeRow;
            return this;
        }

        public EasySheet build() {
            return new EasySheet(sheetName, tempSheetStyles, tempEasyRows, freezeCol, freezeRow);
        }
    }
}
