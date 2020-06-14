package com.boot.app.excel;

import java.io.IOException;
import java.lang.reflect.Field;
import java.util.List;
import java.util.Map;
import java.util.function.Consumer;
import java.util.stream.Stream;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.http.ContentDisposition;
import org.springframework.http.HttpHeaders;
import org.springframework.util.StopWatch;
import org.springframework.web.servlet.view.document.AbstractXlsxStreamingView;

import lombok.extern.slf4j.Slf4j;

@Slf4j
public class ExcelView<T> extends AbstractXlsxStreamingView {
    
    public static final String EXCEL_MODEL = "excelModel";
    
    @Override
    protected void buildExcelDocument(Map<String, Object> model, Workbook workbook, 
            HttpServletRequest request, HttpServletResponse response) throws Exception {
        @SuppressWarnings("unchecked")
        ExcelModel<T> excelModel = (ExcelModel<T>)model.get(EXCEL_MODEL);
        setFileName(response, excelModel);
        Sheet sheet = workbook.createSheet();
        customize(workbook, excelModel);
        setRows(sheet, excelModel);
    }
    private void setFileName(HttpServletResponse response, ExcelModel<T> excelModel) {
        String fileName = excelModel.getFileName();
        String contentDisposition = ContentDisposition.builder("attachment").filename(fileName).build().toString();
        response.setHeader(HttpHeaders.CONTENT_DISPOSITION, contentDisposition);
    }
    private void customize(Workbook workbook, ExcelModel<T> excelModel) {
        Consumer<Workbook> customizer = excelModel.getCustomizer();
        if (customizer != null) {
            customizer.accept(workbook);
        }
    }
    private void setRows(Sheet sheet, ExcelModel<T> excelModel) {
        String[] keys = excelModel.getHeaderKeys();
        String[] names = excelModel.getHeaderNames();
        int len = keys.length;
        setHeaderRow(sheet, names, len);
        List<T> rows = excelModel.getRows();
        if (rows.isEmpty()) {
            return;
        }
        setBodyRows(sheet, keys, rows);
    }

    protected void setHeaderRow(Sheet sheet, String[] names, int len) {
        Row headerRow = sheet.createRow(0);
        CellStyle centerAlign = sheet.getWorkbook().createCellStyle();
        centerAlign.setAlignment(HorizontalAlignment.CENTER);
        for(int j=0; j<len; j++) {
            Cell headerCell = headerRow.createCell(j);
            headerCell.setCellValue(names[j]);
            headerCell.setCellStyle(centerAlign);
        }
    }
    
    protected void setBodyRows(Sheet sheet, String[] keys, List<T> rows) {
        Class<?> clz = rows.get(0).getClass();
        Field[] fields = Stream.of(keys).map(key -> getField(clz, key)).toArray(Field[]::new);
        int i = 1;
        for(T each : rows) {
            Row row = sheet.createRow(i);
            int j = 0;
            for(Field  field : fields) {
                Cell cell = row.createCell(j);
                cell.setCellValue(getFieldAsString(each, field));
                j += 1;
            }
            i += 1;
        }
    }    
    private String getFieldAsString(Object obj, Field field) {
        try {
            return (String)field.get(obj);
        } 
        catch (Exception e) {
            throw new IllegalArgumentException(e);
        }         
    }
    private Field getField(Class<?> clz, String name) {
        try {
            Field field = clz.getDeclaredField(name);
            field.setAccessible(true);
            return field;
        } 
        catch (Exception e) {
            throw new IllegalArgumentException(e);
        }       
    }
    
    @Override
    protected void renderWorkbook(Workbook workbook, HttpServletResponse response) throws IOException {
        StopWatch stopWatch = new StopWatch();
        stopWatch.start("ExcelView");
        super.renderWorkbook(workbook, response);
        stopWatch.stop();
        log.info(stopWatch.prettyPrint());
    }    

}
