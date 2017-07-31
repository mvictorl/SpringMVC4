package com.mvictorl.mvc.excelpdf;

import org.apache.poi.ss.usermodel.*;
import org.springframework.web.servlet.view.document.AbstractXlsView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.List;
import java.util.Map;

public class ExcelDocument extends AbstractXlsView {
    @Override
    protected void buildExcelDocument(
            Map<String, Object> map,
            Workbook workbook,
            HttpServletRequest httpServletRequest,
            HttpServletResponse httpServletResponse) throws Exception {
        Sheet excelSheet = workbook.createSheet("Simple excel example.");
        httpServletResponse.setHeader("Content-Disposition", "attachment; filename=excelDocument.xls");

        Font font = workbook.createFont();
        font.setFontName("Arial");
        font.setBold(true);
        font.setColor(IndexedColors.WHITE.index);

        CellStyle styleHeader = workbook.createCellStyle();
        styleHeader.setFillForegroundColor(IndexedColors.BLUE.index);
        styleHeader.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        styleHeader.setFont(font);

        setExcelHeader(excelSheet, styleHeader);

        // Get data from model
        List<Cat> cats = (List<Cat>) map.get("modelObject");
        int rowCount = 1;
        for (Cat cat : cats) {
            Row row = excelSheet.createRow(rowCount++);
            row.createCell(0).setCellValue(cat.getName());
            row.createCell(1).setCellValue(cat.getWeight());
            row.createCell(2).setCellValue(cat.getColor());
        }
    }

    private void setExcelHeader(Sheet excelSheet, CellStyle styleHeader) {
        // Set Excel Header names
        Row header = excelSheet.createRow(0);
        header.createCell(0).setCellValue("Name");
        header.getCell(0).setCellStyle(styleHeader);
        header.createCell(1).setCellValue("Wieght");
        header.getCell(1).setCellStyle(styleHeader);
        header.createCell(2).setCellValue("Color");
        header.getCell(2).setCellStyle(styleHeader);
    }
}
