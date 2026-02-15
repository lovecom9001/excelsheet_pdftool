package com.example.excel;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Path;

public class ExcelSheetPdfApp {

    public static void main(String[] args) throws Exception {
        if (args.length != 6) {
            System.out.println("Usage: java -jar excel-sheet-pdf-tool.jar <input.xlsx> <sheetName> <cellRef> <value> <output.xlsx> <output.pdf>");
            System.out.println("Example: java -jar excel-sheet-pdf-tool.jar input.xlsx Sheet1 B3 hello output.xlsx sheet1.pdf");
            return;
        }

        Path inputPath = Path.of(args[0]);
        String sheetName = args[1];
        String cellRef = args[2];
        String value = args[3];
        Path outputExcelPath = Path.of(args[4]);
        Path outputPdfPath = Path.of(args[5]);

        process(inputPath, sheetName, cellRef, value, outputExcelPath, outputPdfPath);
        System.out.println("완료: 엑셀 수정 및 PDF 생성");
    }

    public static void process(Path inputPath,
                               String sheetName,
                               String cellRef,
                               String value,
                               Path outputExcelPath,
                               Path outputPdfPath) throws IOException {
        try (FileInputStream fis = new FileInputStream(inputPath.toFile());
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheet(sheetName);
            if (sheet == null) {
                throw new IllegalArgumentException("시트를 찾을 수 없습니다: " + sheetName);
            }

            CellReferencePos pos = CellReferencePos.parse(cellRef);
            Row row = sheet.getRow(pos.row());
            if (row == null) {
                row = sheet.createRow(pos.row());
            }

            Cell cell = row.getCell(pos.col());
            if (cell == null) {
                cell = row.createCell(pos.col());
            }

            cell.setCellValue(value);

            try (FileOutputStream fos = new FileOutputStream(outputExcelPath.toFile())) {
                workbook.write(fos);
            }

            exportSheetToPdf(sheet, outputPdfPath);
        }
    }

    private static void exportSheetToPdf(Sheet sheet, Path outputPdfPath) throws IOException {
        try (PDDocument document = new PDDocument()) {
            PDPage page = new PDPage(PDRectangle.A4);
            document.addPage(page);

            float margin = 40;
            float yStart = page.getMediaBox().getHeight() - margin;
            float y = yStart;
            float rowHeight = 16;

            try (PDPageContentStream contentStream = new PDPageContentStream(document, page)) {
                contentStream.setFont(PDType1Font.HELVETICA, 10);

                for (Row row : sheet) {
                    if (y < margin) {
                        contentStream.close();
                        page = new PDPage(PDRectangle.A4);
                        document.addPage(page);
                        y = yStart;
                    }

                    StringBuilder line = new StringBuilder();
                    short lastCellNum = row.getLastCellNum();
                    if (lastCellNum < 0) {
                        y -= rowHeight;
                        continue;
                    }

                    for (int c = 0; c < lastCellNum; c++) {
                        Cell cell = row.getCell(c, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                        String text = cell == null ? "" : getCellText(cell);
                        line.append(text);
                        if (c < lastCellNum - 1) {
                            line.append(" | ");
                        }
                    }

                    contentStream.beginText();
                    contentStream.newLineAtOffset(margin, y);
                    contentStream.showText(trimToPdfSafeText(line.toString()));
                    contentStream.endText();

                    y -= rowHeight;
                }
            }

            document.save(outputPdfPath.toFile());
        }
    }

    private static String getCellText(Cell cell) {
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> String.valueOf(cell.getNumericCellValue());
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            case FORMULA -> cell.getCellFormula();
            case BLANK, _NONE, ERROR -> "";
        };
    }

    private static String trimToPdfSafeText(String text) {
        String sanitized = text.replace('\n', ' ').replace('\r', ' ');
        return sanitized.length() > 180 ? sanitized.substring(0, 180) : sanitized;
    }

    private record CellReferencePos(int row, int col) {
        static CellReferencePos parse(String cellRef) {
            int idx = 0;
            while (idx < cellRef.length() && Character.isLetter(cellRef.charAt(idx))) {
                idx++;
            }
            if (idx == 0 || idx == cellRef.length()) {
                throw new IllegalArgumentException("잘못된 셀 주소입니다. 예: B3");
            }

            String colPart = cellRef.substring(0, idx).toUpperCase();
            String rowPart = cellRef.substring(idx);

            int col = 0;
            for (char ch : colPart.toCharArray()) {
                col = col * 26 + (ch - 'A' + 1);
            }
            col -= 1;

            int row = Integer.parseInt(rowPart) - 1;
            if (row < 0 || col < 0) {
                throw new IllegalArgumentException("행/열 번호는 1 이상이어야 합니다.");
            }
            return new CellReferencePos(row, col);
        }
    }
}
