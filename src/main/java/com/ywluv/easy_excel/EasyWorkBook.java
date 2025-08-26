package com.ywluv.easy_excel;


import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.util.StringUtils;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 * packageName    : com.ywluv.easyExcel
 * fileName       : EasyWorkBook.java
 * author         : MYH
 * date           : 2025-08-25
 * description    :
 * ===========================================================
 * DATE              AUTHOR             NOTE
 * -----------------------------------------------------------
 * 2025-08-25        MYH       최초 생성
 */
@Slf4j
public class EasyWorkBook extends EasyStyleResolver{
    private String workbookName;
    private final List<EasySheet> easySheets;

    /**
     * Instantiates a new EasyWorkBook
     * author : MYH
     * description :
     *
     * @param workbookName
     * @param easySheets
     */
    private EasyWorkBook(String workbookName, List<EasySheet> easySheets) {
        this.workbookName = workbookName;
        this.easySheets = Collections.unmodifiableList(easySheets != null ? new ArrayList<>(easySheets) : new ArrayList<>());
    }

    public static Builder builder() {return new Builder();}

    public String getWorkbookName() {
        if(StringUtils.hasText(workbookName)){
            return workbookName+".xlsx";
        } else {
            return "ExcelFile"+".xlsx";
        }
    }

    public List<EasySheet> getRows() {
        return easySheets;
    }

    private Workbook toWorkbook(){
        XSSFWorkbook workbook = new XSSFWorkbook();
        int sheetCount = 1;
        for (EasySheet easySheet : easySheets) {
            String sheetName = StringUtils.hasText(easySheet.getSheetName())
                    ? easySheet.getSheetName()
                    : "Sheet" + sheetCount++;
            var sheetStyle = easyStyleToCellStyle(workbook,easySheet.getSheetStyles());
            XSSFSheet sheet = workbook.createSheet(sheetName);
            sheet.createFreezePane(easySheet.getFreezeCol(),easySheet.getFreezeRow());
            int rowNum = 0;
            for (EasyRow easyRow : easySheet.getRows()) {
                XSSFRow row = sheet.createRow(rowNum++);
                var rowStyle = easyStyleToCellStyle(workbook, easyRow.getRowStyles());
                var cells = easyRow.getCells();
                int columnNum = 0;
                for (var easyCell : cells) {
                    XSSFCell cell = row.createCell(columnNum++);
                    var cellStyle = easyStyleToCellStyle(workbook,easyCell.getCellStyles());
                    selectStyle(cell,sheetStyle,rowStyle,cellStyle);
                    cell.setCellValue(easyCell.getValue());
                }
            }
        }

        return workbook;
    }

    private XSSFCell selectStyle(XSSFCell cell, CellStyle sheetStyle, CellStyle rowStyle, CellStyle cellStyle) {
        if (cellStyle != null) {
            cell.setCellStyle(cellStyle);
        } else if (rowStyle != null) {
            cell.setCellStyle(rowStyle);
        } else if (sheetStyle != null) {
            cell.setCellStyle(sheetStyle);
        }
        return cell;
    }

    public ResponseEntity<byte[]> toResponseEntity(){
        Workbook workbook = toWorkbook();
        byte[] excelBytes;

        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            workbook.write(bos);
            excelBytes = bos.toByteArray(); //
        } catch (IOException e) {
            throw new IllegalStateException("Excel Download Fail");
        } finally {
            try {
                workbook.close(); //
            } catch (IOException e) {
                throw new IllegalStateException("Excel Download Fail");
            }
        }
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION,
                        "attachment;filename=" + URLEncoder.encode(getWorkbookName(), StandardCharsets.UTF_8))
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .contentLength(excelBytes.length)
                .body(excelBytes);

    }

    public static class Builder {
        private String workbookName;
        private List<EasySheet> tempEasySheets = new ArrayList<>();

        public Builder workbookName(String workbookName) {
            this.workbookName = workbookName;
            return this;
        }

        public Builder addSheet(EasySheet easySheet) {
            if (easySheet != null) {
                tempEasySheets.add(easySheet);
            }
            return this;
        }

        public Builder addSheet(EasySheet easySheet, Integer index) {
            if (easySheet != null) {
                int pos;
                if (index == null || index < 0) {
                    pos = 0;
                } else if (index > tempEasySheets.size()) {
                    pos = tempEasySheets.size();
                } else {
                    pos = index;
                }
                tempEasySheets.add(pos, easySheet);
            }
            return this;
        }

        public Builder addSheetAtEnd(List<EasySheet> easySheets) {
            if (easySheets != null && !easySheets.isEmpty()) {
                tempEasySheets.addAll(easySheets);
            }
            return this;
        }

        public Builder addRowsAtStart(List<EasySheet> easySheets) {
            if (easySheets != null && !easySheets.isEmpty()) {
                tempEasySheets.addAll(0, easySheets);
            }
            return this;
        }

        public EasyWorkBook build() {
            return new EasyWorkBook(workbookName,tempEasySheets);
        }
    }



}
