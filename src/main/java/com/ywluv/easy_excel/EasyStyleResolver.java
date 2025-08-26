package com.ywluv.easy_excel;


import com.ywluv.easy_excel.enums.EasyCellStyleEnum;
import org.apache.poi.ss.usermodel.*;
import java.util.Set;

/**
 * packageName    : com.ywluv.easyExcel
 * fileName       : EasyStyleResolver.java
 * author         : MYH
 * date           : 2025-08-25
 * description    :
 * ===========================================================
 * DATE              AUTHOR             NOTE
 * -----------------------------------------------------------
 * 2025-08-25        MYH       최초 생성
 */
public abstract class EasyStyleResolver {

    protected static CellStyle easyStyleToCellStyle(Workbook workbook, Set<EasyCellStyleEnum> styleEnums){
        if (styleEnums == null || styleEnums.isEmpty()) {return null;}
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        for (var s : styleEnums) {
            switch (s) {
                case DEFAULT -> {
                    style.setAlignment(HorizontalAlignment.CENTER);
                    style.setVerticalAlignment(VerticalAlignment.CENTER);
                }
                case TITLE -> {
                    style.setAlignment(HorizontalAlignment.CENTER);
                    style.setVerticalAlignment(VerticalAlignment.CENTER);
                    font.setBold(true);
                    style.setAlignment(HorizontalAlignment.CENTER);
                    style.setVerticalAlignment(VerticalAlignment.CENTER);
                    style.setBorderTop(BorderStyle.THIN);
                    style.setBorderBottom(BorderStyle.THIN);
                    style.setBorderLeft(BorderStyle.THIN);
                    style.setBorderRight(BorderStyle.THIN);
                }
                case FONT_BOLD -> font.setBold(true);
                case ALIGNMENT -> style.setAlignment(HorizontalAlignment.CENTER);
                case VERTICAL_ALIGNMENT -> style.setVerticalAlignment(VerticalAlignment.CENTER);
                case BORDER_ALL -> {
                    style.setBorderTop(BorderStyle.THIN);
                    style.setBorderBottom(BorderStyle.THIN);
                    style.setBorderLeft(BorderStyle.THIN);
                    style.setBorderRight(BorderStyle.THIN);
                }
                case BORDER_TOP -> style.setBorderTop(BorderStyle.THIN);
                case BORDER_BOTTOM -> style.setBorderBottom(BorderStyle.THIN);
                case BORDER_RIGHT -> style.setBorderRight(BorderStyle.THIN);
                case BORDER_LEFT -> style.setBorderLeft(BorderStyle.THIN);
            }
        }
        style.setFont(font);
        return style;
    }
}
