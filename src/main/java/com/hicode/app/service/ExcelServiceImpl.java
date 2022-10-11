package com.hicode.app.service;

import com.hicode.app.dto.ProductDTO;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.*;

@Service
public class ExcelServiceImpl implements ExcelService {

    public static final int COLUMN_INDEX_ID = 0;
    public static final int COLUMN_INDEX_NAME = 1;
    public static final int COLUMN_INDEX_PRICE = 2;
    public static final int COLUMN_INDEX_QUALITY = 3;
    public static final int COLUMN_INDEX_CREATEDATE = 4;
    private static CellStyle cellStyleFormatNumber = null;
    @Override
    public List<ProductDTO> readFile(MultipartFile file) throws IOException {
        List<ProductDTO> listBooks = new ArrayList<>();
        //get workbook
        Workbook workbook = getWorkbook(file.getInputStream(), file.getOriginalFilename());
        //get sheet
        Sheet sheet = workbook.getSheetAt(0);
        //getAll rows
        Iterator<Row> iterator = sheet.iterator();
        while (iterator.hasNext()) {
            Row nextRow = iterator.next();
            if (nextRow.getRowNum() == 0) {
                //ignore header
                continue;
            }
            //getAll cells
            Iterator<Cell> cellIterator = nextRow.cellIterator();
            //Read cells and set value for book object
            ProductDTO productDTO = new ProductDTO();
            while (cellIterator.hasNext()) {
                //Read cell
                Cell cell = cellIterator.next();
                Object cellValue = getCellValue(cell);
                if (cellValue == null || cellValue.toString().isEmpty()) {
                    continue;
                }
                //Set value;
                int columnIndex = cell.getColumnIndex();
                switch (columnIndex) {
                    case COLUMN_INDEX_ID:
                        productDTO.setId(cell.getStringCellValue());
                        break;
                    case COLUMN_INDEX_NAME:
                        productDTO.setName(cell.getStringCellValue());
                        break;
                    case COLUMN_INDEX_PRICE:
                        productDTO.setPrice((cell.getNumericCellValue()));

                        break;
                    case COLUMN_INDEX_QUALITY:
                        productDTO.setQuality((int) cell.getNumericCellValue());
                        break;
                    case COLUMN_INDEX_CREATEDATE:
                        productDTO.setCreateDate(cell.getDateCellValue());
                        break;
                    default:
                        break;
                }
            }
            listBooks.add(productDTO);

        }
        workbook.close();

        return listBooks;
    }

    private Object getCellValue(Cell cell) {
        CellType cellType = cell.getCellTypeEnum();
        Object cellValue = null;
        switch (cellType) {
            case BOOLEAN:
                cellValue = cell.getBooleanCellValue();
                break;
            case FORMULA:
                Workbook workbook = cell.getSheet().getWorkbook();
                FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                cellValue = evaluator.evaluate(cell).getStringValue();
                break;
            case NUMERIC:
                cellValue = cell.getNumericCellValue();
                break;
            case STRING:
                cellValue = cell.getStringCellValue();
                break;
            case _NONE:
            case BLANK:
            case ERROR:
                break;
            default:
                break;
        }


        return cellValue;
    }

    private Workbook getWorkbook(InputStream inputStream, String filename) throws IOException {
        Workbook workbook = null;
        if (filename.endsWith("xlsx")) {
            workbook = new XSSFWorkbook(inputStream);
        } else if (filename.endsWith("xls")) {
            workbook = new HSSFWorkbook(inputStream);
        } else throw new IllegalArgumentException("The specified file is not Excel file");

        return workbook;
    }

    @Override
    public void saveDatatoDB() {

    }

    private static final List<ProductDTO> FAKE_DATA = Arrays.asList(
            new ProductDTO("P1", "SamSung", 123, 12, new Date()),
            new ProductDTO("P1", "SamSung", 123, 12, new Date()),
            new ProductDTO("P1", "SamSung", 123, 12, new Date()),
            new ProductDTO("P1", "SamSung", 123, 12, new Date()),
            new ProductDTO("P1", "SamSung", 123, 12, new Date())

    );

    @Override
    public ByteArrayInputStream downloadFile(String filename) throws IOException {
        Workbook workbook = getWorkbook(filename);

        //Create sheet
        Sheet sheet = workbook.createSheet("Products");

        int rowIndex = 0;

        // write header
        writeHeader(sheet, rowIndex);

        rowIndex++;

        for (ProductDTO productDTO : FAKE_DATA
        ) {
            //create row;
            Row row = sheet.createRow(rowIndex);
            //write data on row
            writeBook(productDTO, row);
            rowIndex++;
        }
        //Auto resize column witdth
        int numberOfColumn = sheet.getRow(0).getPhysicalNumberOfCells();
        autosizeColumn(sheet, numberOfColumn);
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        workbook.write(out);
        return new ByteArrayInputStream(out.toByteArray());
    }



    private void autosizeColumn(Sheet sheet, int numberOfColumn) {
        for (int columnIndex = 0; columnIndex < numberOfColumn; columnIndex++) {
            sheet.autoSizeColumn(columnIndex);
        }
    }

    private void writeBook(ProductDTO productDTO, Row row) {
        if (cellStyleFormatNumber == null) {
            // Format number
            short format = (short)BuiltinFormats.getBuiltinFormat("#,##0");
            // DataFormat df = workbook.createDataFormat();
            // short format = df.getFormat("#,##0");

            //Create CellStyle
            Workbook workbook = row.getSheet().getWorkbook();
            cellStyleFormatNumber = workbook.createCellStyle();
            cellStyleFormatNumber.setDataFormat(format);
        }
        Cell cell = row.createCell(COLUMN_INDEX_ID);
        cell.setCellValue(productDTO.getId());

        cell = row.createCell(COLUMN_INDEX_NAME);
        cell.setCellValue(productDTO.getName());

        cell = row.createCell(COLUMN_INDEX_PRICE);
        cell.setCellValue(productDTO.getPrice());
        cell.setCellStyle(cellStyleFormatNumber);

        cell = row.createCell(COLUMN_INDEX_QUALITY);
        cell.setCellValue(productDTO.getQuality());

        cell = row.createCell(COLUMN_INDEX_CREATEDATE);
        cell.setCellValue(productDTO.getCreateDate());
    }

    private void writeHeader(Sheet sheet, int rowIndex) {
        //create CellStyle
        CellStyle cellStyle = createStyleForHeader(sheet);

        //create row
        Row row = sheet.createRow(rowIndex);
        //create cell
        Cell cell = row.createCell(COLUMN_INDEX_ID);
        cell.setCellStyle(cellStyle);
        cell.setCellValue("Id");

        cell = row.createCell(COLUMN_INDEX_NAME);
        cell.setCellStyle(cellStyle);
        cell.setCellValue("Name");

        cell = row.createCell(COLUMN_INDEX_PRICE);
        cell.setCellStyle(cellStyle);
        cell.setCellValue("Price");

        cell = row.createCell(COLUMN_INDEX_QUALITY);
        cell.setCellStyle(cellStyle);
        cell.setCellValue("Quality");

        cell = row.createCell(COLUMN_INDEX_CREATEDATE);
        cell.setCellStyle(cellStyle);
        cell.setCellValue("Create Date");
    }

    private CellStyle createStyleForHeader(Sheet sheet) {
        //Create font
        Font font = sheet.getWorkbook().createFont();
        font.setFontName("Times New Roman");
        font.setBold(true);
        font.setFontHeightInPoints((short) 14);
        font.setColor(IndexedColors.WHITE.getIndex());
        //create style
        CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
        cellStyle.setFont(font);
        cellStyle.setFillForegroundColor(IndexedColors.BLACK.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        return cellStyle;
    }

    private Workbook getWorkbook(String filename) {
        Workbook workbook = null;
        if (filename.endsWith("xlsx")) {
            workbook = new XSSFWorkbook();
        } else if (filename.endsWith("xls")) {
            workbook = new HSSFWorkbook();
        } else {
            throw new IllegalStateException("The specified file is not Excel file");
        }
        return workbook;
    }
}
